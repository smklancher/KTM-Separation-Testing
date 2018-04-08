'#Reference {3F4DACA7-160D-11D2-A8E9-00104B365C9F}#5.5#0#C:\Windows\SysWOW64\vbscript.dll\3#Microsoft VBScript Regular Expressions 5.5#VBScript_RegExp_55
'#Reference {420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\SysWOW64\scrrun.dll#Microsoft Scripting Runtime#Scripting
'#Language "WWB-COM"
Option Explicit


Type Param
   ' Represents a parameter to a function as parsed by regex
   Name As String
   ParamType As String
   Array As Boolean
   OptionalParam As Boolean
   DefaultValue As String
End Type

Type ScriptFunction
   ' Represents a script function parsed by regex
   Name As String
   IsSub As Boolean
   Params() As Param
   StartIndex As Long
   EndIndex As Long
   Content As String
   ClassName As String
   ReturnType As String
   Suspect As Boolean
   StringTag As String
End Type



'KTM Script Logging Framework
'2013-09-09 - Splitting batch logs by PID is now optional, default to off (set constant PROCESS_ID_IN_BATCHLOG_FILENAME)
'2013-08-08 - Design-time only issue in PB 6.0 fixed: When ParentFolder not known, doc/page not recorded.
'2013-05-16 - Split batch log by process ID.  Log process ID on error.
'2013-01-16 - To simplify code, introduced dependency on Scripting.FileSystemObject.
'           - Introduced dependency on Wscript.Shell to get correct path regardless of version or language of the OS.
'           - Fix for negative error numbers not logged
'2012-12-16 - Delimiter was missing, causing Windows user and module to be combined

'CONFIGURABLE LOGGING CONSTANTS

'This message will be prepended to the msgbox shown to a user if there is a script error
Public Const USER_ERROR_MSG As String = "An error has occurred in the KTM project script.  " & _
   "If this message needs to be reported to a system administrator, please take a " & _
   "screenshot with this message showing and explain what actions were taken immediately " & _
   "before the error occurred."
Public Const LOG_FILENAME As String = "_KTM_Script_Log.log"
Public Const BATCH_LOG_FILENAME As String = "_KTM_Script_Batch.log"
Public Const DESIGN_LOG_FILENAME As String = "_KTM_Script_Design.log"

Public BATCH_LOG_FULLPATH As String

'for readability messages are logged on a newline after metadata, set this to true to log to a single line
Public Const LOG_SINGLE_LINE As Boolean = False


'NONCONFIGURABLE LOGGING CONSTANTS
Public Const LOCAL_LOG As Boolean = True
Public Const BATCH_LOG As Boolean = False
Public Const IGNORE_CURRENT_FUNCTION As Integer = 1
Public Const FORCE_ERROR As Boolean = True
Public Const ONLY_ON_ERROR As Boolean = False
Public Const SUPPRESS_MSGBOX As Boolean = True


'SMK 2013-05-16 Use to get process ID in InitializeBatch
'SMK 2013-09-09 Make per process log optional, default to off
Declare Function GetCurrentProcessId Lib "kernel32" Alias "GetCurrentProcessId" () As Long
Public PROCESS_ID As Long
Public Const PROCESS_ID_IN_BATCHLOG_FILENAME As Boolean = False


'GLOBAL LOGGING VARIABLES
'This will be changed to true if it looks like we are in a Thin Client module
Public THIN_CLIENT As Boolean

Public BATCH_IMAGE_LOGS As String
Public CAPTURE_LOCAL_LOGS As String

Public BATCH_CLASS As String
Public BATCH_NAME As String
Public BATCH_ID As Long
Public BATCH_ID_HEX As String

'to support KTM 5.5 features
Public BATCH_USERID As String
Public BATCH_USERNAME As String
Public BATCH_WINDOWSUSERNAME As String
Public BATCH_USERSTRING As String 'combination of the previous three




'======== START LOGGING CODE ========

'Initialize Capture/runtime info.  Call from Application_InitializeBatch
Public Sub Logging_InitializeBatch(ByVal pXRootFolder As CscXFolder)
   'SMK 2013-05-16 - Determine process ID
   On Error Resume Next
   PROCESS_ID=GetCurrentProcessId()

   On Error GoTo CouldNotCreate
   'assume the batch log folder does not exist
   Dim LogFolderExists As Boolean
   LogFolderExists = False

   'these items are only set by Capture at runtime.  if any are set,
   ' then we are at runtime and they are all set
   If pXRootFolder.XValues.ItemExists("AC_BATCH_CLASS_NAME") Then
      'Set batchname, batchid, batch class
      BATCH_CLASS = pXRootFolder.XValues.ItemByName("AC_BATCH_CLASS_NAME").Value
      BATCH_NAME = pXRootFolder.XValues.ItemByName("AC_BATCH_NAME").Value
      BATCH_ID = CLng(pXRootFolder.XValues.ItemByName("AC_EXTERNAL_BATCHID").Value)
      BATCH_ID_HEX = Hex(BATCH_ID)

      'pad hex ID
      BATCH_ID_HEX = Right("00000000", 8 - Len(BATCH_ID_HEX)) & BATCH_ID_HEX

      'These items are only present in KTM 5.5
      If pXRootFolder.XValues.ItemExists("AC_BATCH_WINDOWSUSERNAME") Then
         BATCH_WINDOWSUSERNAME=pXRootFolder.XValues.ItemByName("AC_BATCH_WINDOWSUSERNAME").Value
         BATCH_USERID = pXRootFolder.XValues.ItemByName("AC_BATCH_USERID").Value
         BATCH_USERNAME = pXRootFolder.XValues.ItemByName("AC_BATCH_USERNAME").Value

         'if user profiles are off these will all be the same
         If BATCH_WINDOWSUSERNAME = BATCH_USERID And BATCH_USERID = BATCH_USERNAME Then
            'user profiles is off so only use the windows user
         'SMK 2012-12-16 Delimiter was missing, causing Windows user and module to be combined
            BATCH_USERSTRING = BATCH_WINDOWSUSERNAME & " -- "
         Else
            'user profiles is on, so use all
            BATCH_USERSTRING = BATCH_WINDOWSUSERNAME & ", " & BATCH_USERNAME & _
               " (" & BATCH_USERID & ") -- "
         End If
      End If

      'set the batch logging path
      BATCH_IMAGE_LOGS = pXRootFolder.XValues.ItemByName("AC_IMAGE_DIRECTORY").Value & _
         "\" & BATCH_ID_HEX & "\Log\"

      'SMK 2013-01-16 - To simplify code, introduced dependency on Scripting.FileSystemObject.

      'To use an early bound object(FileSystemObject), add a reference to "Microsoft Scripting Runtime"
      ' C:\Windows\System32\scrrun.dll (C:\Windows\SysWOW64\scrrun.dll)
      ' Otherwise late bound object will be created via CreateObject("Scripting.FileSystemObject")
      Dim fso As Object
      Set fso = CreateObject("Scripting.FileSystemObject")
      'Dim fso As FileSystemObject

      If Not fso.FolderExists(BATCH_IMAGE_LOGS) Then
         fso.CreateFolder(BATCH_IMAGE_LOGS)
      End If
      LogFolderExists=True

      'if creating the folder causes an error then FolderExists is still false
      CouldNotCreate:
      Err.Clear()
      On Error GoTo catch


      'if the folder still doesn't exist after trying to create, just use image path
      If Not LogFolderExists Then
         'we prefer to log to the "Log" folder along with the interactive modules,
         '  but if there is a problem, use the image path itself
         BATCH_IMAGE_LOGS = pXRootFolder.XValues.ItemByName("AC_IMAGE_DIRECTORY").Value & _
            "\" & BATCH_ID_HEX & "\"
      End If

      'Set Process ID as part of the batch log path if needed
      If PROCESS_ID_IN_BATCHLOG_FILENAME Then
         'SMK 2013-05-16 - Separate log per process, ID set in InitializeBatch
         'SMK 2013-09-09 - Made optional, default to off
         BATCH_LOG_FULLPATH = BATCH_IMAGE_LOGS & BATCH_ID_HEX & "_" & PROCESS_ID & BATCH_LOG_FILENAME
         'otherwise leave blank and it gets set as needed in ScriptLog
      End If
   End If

   catch:
   Set fso=Nothing
End Sub


'log initial information about the batch.  Call from Batch_Open
Public Sub Logging_BatchOpen(ByVal pXRootFolder As CscXFolder)
   On Error GoTo catch

   'the project file is copied on publish and retains its original modified date
   Dim ProjectLastSave As Date


   ProjectLastSave = FileDateTime(Project.FileName)

   'We can only get the batch class publish date if "Copy project during publish" is used
   '  otherwise we will just get the project path
   Dim BatchClassPublishOrProjectPath As String

   'if the "Copy project during publish" is checked, it will be located within PubTypes\Custom
   If InStr(1, Project.FileName, "PubTypes\Custom") > 0 Then
      'with "Copy project during publish" the folder containing the project is created
      '   (thus dated) while publishing
      Dim ProjectFolder As String
      ProjectFolder = Mid(Project.FileName, 1, InStrRev(Project.FileName, "\") - 1)

      BatchClassPublishOrProjectPath = "published " & CStr(FileDateTime(ProjectFolder))
   Else
      'without "copy project during publish" the folder could have any date
      '  and the project could be anywhere, so just get the project path
      BatchClassPublishOrProjectPath = Project.FileName
   End If

   'log basics like batch name, class, id, machine name, project save date
   ScriptLog("Opening Batch """ & BATCH_NAME & """ (" & BATCH_ID & "/" & _
      BATCH_ID_HEX & ") -- " & BATCH_USERSTRING & Environ("ComputerName") & vbNewLine & _
      "Batch " & BATCH_ID & "/" & BATCH_ID_HEX & ": Batch Class """ & BATCH_CLASS & _
      """ (" & BatchClassPublishOrProjectPath & ", project saved " & CStr(ProjectLastSave) _
      & ")", LOCAL_LOG)

   Exit Sub

   'if there is an error, log it and try to keep going
   catch:
   ErrorLog(Err, "", Nothing, pXRootFolder, 0, ONLY_ON_ERROR, SUPPRESS_MSGBOX)
   Resume Next
End Sub



'Find and return the Capture\Local\Logs directory
'   caching the result to global variable CAPTURE_LOCAL_LOGS
Public Function Logging_CaptureLocalLogs() As String

   'If CAPTURE_LOCAL_LOGS is already set, just return it
   If CAPTURE_LOCAL_LOGS <> "" Then
      Logging_CaptureLocalLogs = CAPTURE_LOCAL_LOGS
      Exit Function

   Else
      On Error GoTo CouldNotCreate

      'SMK 2013-01-16 - Introduced dependency on Wscript.Shell to read "Local" path from registry.
      '                 Unlike previous method of manipulating environment path variables, this provides
      '                 the correct path regardless of version or language of the OS.

      'To use an early bound object (WshShell), add a reference to "Windows Script Host Object Model"
      ' C:\Windows\System32\wshom.ocx (C:\Windows\SysWOW64\wshom.ocx)
      ' Otherwise late bound object will be created via CreateObject("Wscript.Shell")
      Dim wsh As Object
      Set wsh = CreateObject("Wscript.Shell")
      'Dim wsh As New WshShell

      'The same registry location works on 32 or 64 bit OS (Windows redirect registry access from 32-bit apps to Wow6432Node as needed)
      CAPTURE_LOCAL_LOGS=wsh.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Kofax Image Products\Ascent Capture\3.0\LocalPath") & "\Logs\"

      'assume the batch log folder does not exist
      Dim LogFolderExists As Boolean
      LogFolderExists = False

      'SMK 2013-01-16 - To simplify code, introduced dependency on Scripting.FileSystemObject.

      'To use an early bound object(FileSystemObject), add a reference to "Microsoft Scripting Runtime"
      ' C:\Windows\System32\scrrun.dll (C:\Windows\SysWOW64\scrrun.dll)
      ' Otherwise late bound object will be created via CreateObject("Scripting.FileSystemObject")
      Dim fso As Object
      Set fso = CreateObject("Scripting.FileSystemObject")
      'Dim fso As FileSystemObject

      If Not fso.FolderExists(CAPTURE_LOCAL_LOGS) Then
         fso.CreateFolder(CAPTURE_LOCAL_LOGS)
      End If
      LogFolderExists=True

      'if creating the folder causes an error then FolderExists is still false
      CouldNotCreate:
      Err.Clear()
      On Error GoTo catch

      'if the folder still doesn't exist after trying to create, just use image path
      If Not LogFolderExists Then
         'if there is a problem getting the Capture local logs folder,
         '   log to the system temp folder
         CAPTURE_LOCAL_LOGS = Environ("Temp") & "\"
      End If

      Logging_CaptureLocalLogs = CAPTURE_LOCAL_LOGS
   End If

   catch:
      Set wsh=Nothing
      Set fso=Nothing
End Function



'primary logging function
Public Sub ScriptLog(ByVal msg As String, Optional ByVal AddToLocalLog As Boolean = False, _
   Optional ByVal pXDoc As CscXDocument = Nothing, _
   Optional ByVal pXFolder As CscXFolder = Nothing, _
   Optional ByVal PageNum As Integer = 0, _
   Optional ByVal ExtraDepth As Integer = 0)

   On Error GoTo catch

   'for readability messages are logged on a newline after metadata,
   '   but this is a matter of preference
   If Not LOG_SINGLE_LINE Then
      msg = Replace(vbNewLine & msg, vbNewLine, vbNewLine & vbTab)
   End If

   'refer to the WinWrap documentation for "CallersLine Function" for
   '   an explaination of the Depth parameter
   Dim Caller As String
   Caller = CallersLine(ExtraDepth)

   'In addition to the date/time, Timer makes it easy to see the number
   '  of seconds/hundredths of seconds between events
   '  output of Timer is also padded for readability in the log
   Dim DateString As String
   DateString = Now & " (" & Format(Timer, "00000.00") & ") -- "

   'if we have an folder/xdoc/page (optional) we can log which document for context
   Dim WhichDoc As String
   WhichDoc = Logging_IdentifyFolderDocPage(pXFolder, pXDoc, PageNum)

   If WhichDoc <> "" Then
      WhichDoc = WhichDoc & " -- "
   End If

   'current module, class/function/line logged from
   Dim ModuleAndFunction As String
   ModuleAndFunction = Logging_ExecutionModeString() & " -- " & Logging_StackLine(Caller) & " "

   ' Always print a shorter message to the intermediate pane of the script window
   ' Previously this was only done in "Design" modes, but this meant it excluded testing runtime events
   Debug.Print(Format(Timer, "00000.00") & " " & WhichDoc & ModuleAndFunction & msg)

   'check if we are in design or runtime
   If Project.ScriptExecutionMode = CscScriptModeServerDesign Or _
      Project.ScriptExecutionMode = CscScriptModeValidationDesign Or _
      Project.ScriptExecutionMode = CscScriptModeVerificationDesign Then

      'if we are in project builder, there is no "image directory" so just write to local logs
      '  Because Application_InitializeBatch is only called manually in PB,
      '  use Logging_CaptureLocalLogs() directly to ensure the path is set
      Open Logging_CaptureLocalLogs() & Format(Now(), "yyyymmdd") & DESIGN_LOG_FILENAME _
         For Append As #1
         Print #1, DateString & WhichDoc & ModuleAndFunction & msg
      Close #1
   Else
      'In case logging is attempted without or before initialization
      If BATCH_IMAGE_LOGS = "" Then
         BATCH_IMAGE_LOGS = Environ("Temp") & "\"
         BATCH_ID_HEX = "Unknown"
      End If

      'If using the process ID in the filename this will have been set from InitializeBatch
      If BATCH_LOG_FULLPATH = "" Then
         BATCH_LOG_FULLPATH = BATCH_IMAGE_LOGS & BATCH_ID_HEX & BATCH_LOG_FILENAME
      End If

      'During runtime, always log to the batch log
      Open BATCH_LOG_FULLPATH For Append As #1
      'batches are processed from various machines, so each line in the batch log
      '   should specify the machine
         Print #1, DateString & Environ("ComputerName") & " -- "  & BATCH_USERSTRING & _
            WhichDoc & ModuleAndFunction & msg
      Close #1

      'And add to local log if specified
      If AddToLocalLog Then
         Open Logging_CaptureLocalLogs() & Format(Now(), "yyyymmdd") & LOG_FILENAME _
            For Append As #1
         'the machine may be processing different batches/modules/users concurrently,
         '  so each line in the local log should have a batch id, userstring
            Print #1, DateString & BATCH_ID_HEX & " -- "  & BATCH_USERSTRING & _
               WhichDoc & ModuleAndFunction & msg
         Close #1
      End If
   End If

   catch:
End Sub


'Primary error logging function.  First parameter should always be "Err".
Public Sub ErrorLog(ByVal E As ErrObject, Optional ByVal ExtraInfo As String = "", _
   Optional ByVal pXDoc As CscXDocument = Nothing, _
   Optional ByVal pXFolder As CscXFolder = Nothing, _
   Optional ByVal PageNum As Integer = 0, _
   Optional ByVal ForceError As Boolean = False, _
   Optional ByVal SuppressMsgBox As Boolean = False)

   'checking if there is an error here means it does not need to be
   '   checked before the function is called
   If E = 0 And ForceError = False Then
      Exit Sub
   End If

   'ErrorMessage will be displayed To user In interactive modules
   Dim ErrorMessage As String
   ErrorMessage = "[Error] PID " & PROCESS_ID & " - " 'SMK 2013-05-16 include process ID in error

   'FIX SMK 2013-01-16 include negative numbers
   If E <> 0 Then
      ErrorMessage = ErrorMessage & E.Number & " - " & E.Description
   End If

   If ExtraInfo <> "" Then
      ErrorMessage = ErrorMessage & "  " & ExtraInfo
   End If


   'when the error handler is set it clears the error, so we must finish with the e param first
   E.Clear()
   On Error GoTo catch


   'get stack trace
   Dim Stack As String

   '1 extra depth to ignore this current function in the stack
   Stack = Logging_StackTrace(IGNORE_CURRENT_FUNCTION)

   'Add stack trace to the error message
   ErrorMessage = ErrorMessage & vbNewLine & Stack

   'log the error and stacktrace
   ScriptLog(ErrorMessage, LOCAL_LOG, pXDoc, pXFolder, PageNum, IGNORE_CURRENT_FUNCTION)


   'Display to user if not in Server or thin client
   If Project.ScriptExecutionMode <> CscScriptModeServer And Not THIN_CLIENT And _
      Not SuppressMsgBox Then
      'if the message needs to be localized, other languages can be added
      '   as seen in the script help topic:
      'Script Samples | Displaying Translated Error Messages For a Script Validation Method
      Dim LocalizedMessage As String


      Select Case Application.UILanguage
         Case "en-US"  'American English
            LocalizedMessage = USER_ERROR_MSG
         Case Else
            LocalizedMessage = USER_ERROR_MSG
      End Select

      'include info about where we are
      Dim WhichDoc As String
      WhichDoc = Logging_IdentifyFolderDocPage(pXFolder, pXDoc, PageNum)

      'include batch info if it exists
      If BATCH_ID_HEX <> "" Then
         WhichDoc = BATCH_NAME & " (" & BATCH_ID_HEX & "), Batch Class: " & BATCH_CLASS _
            & vbNewLine & WhichDoc
      End If

      MsgBox(LocalizedMessage & vbNewLine & vbNewLine & WhichDoc & vbNewLine & vbNewLine & _
         ErrorMessage, vbCritical, "Script Error")
   End If

   catch:
End Sub


'returns a stacktrace from where ever it is called
Public Function Logging_StackTrace(Optional ByVal ExtraDepth As Integer = 0) As String
   On Error GoTo catch

   Dim i As Integer
   i = ExtraDepth

   Dim CurrentStackLine As String
   CurrentStackLine = CallersLine(i)

   'as long as CallersLine returns something, stacktrace continues
   While CurrentStackLine <> ""
      'get a nicer format for the stack line
      CurrentStackLine = i & ": " & Logging_StackLine(CurrentStackLine) & _
         Mid(CurrentStackLine, InStr(1, CurrentStackLine, "]") + 1)

      'Add current line to the stack trace
      Logging_StackTrace = Logging_StackTrace & CurrentStackLine & vbNewLine

      'increment and try to get the next line (CallersLine returns blank if none)
      i = i + 1
      CurrentStackLine = CallersLine(i)

      'protect against trying to log a large stack
      If i > 10 Then
         Logging_StackTrace = Logging_StackTrace & i & ": ...Stack continues beyond " & _
            i - 1 & " frames..."
         Exit While
      End If
   Wend

   'on error exit
   catch:
End Function


'Returns string from ScriptExecutionMode Enum to indicate which module is running
Public Function Logging_ExecutionModeString() As String
   On Error GoTo catch

   'There is not currently a way to tell if a script is executing in a rich or thin client
   '  This is important because MsgBox cannot be used if we are in a thin client
   '  If a thin client is enabled for the project and we are in that module,
   '  we must assume it is a thin client
   THIN_CLIENT = False

   Select Case Project.ScriptExecutionMode
      Case CscScriptModeServer
         Logging_ExecutionModeString = "Server " & Project.ScriptExecutionInstance
      Case CscScriptModeServerDesign
         Logging_ExecutionModeString = "ServerDesign " & Project.ScriptExecutionInstance
      Case CscScriptModeUnknown
         Logging_ExecutionModeString = "Unknown"
      Case CscScriptModeValidation
         Logging_ExecutionModeString = "Validation " & Project.ScriptExecutionInstance
         If Project.WebBasedValidationEnabled Then
            THIN_CLIENT = True
         End If
      Case CscScriptModeValidationDesign
         Logging_ExecutionModeString = "ValidationDesign " & Project.ScriptExecutionInstance
      Case CscScriptModeVerification
         Logging_ExecutionModeString = "Verification"
         If Project.WebBasedVerificationEnabled Then
            THIN_CLIENT = True
         End If
      Case CscScriptModeVerificationDesign
         Logging_ExecutionModeString = "VerificationDesign"
      Case CscScriptModeDocumentReview
         Logging_ExecutionModeString = "DocumentReview"
         If Project.WebBasedDocumentReviewEnabled Then
            THIN_CLIENT = True
         End If
      Case CscScriptModeCorrection
         Logging_ExecutionModeString = "Correction"
         If Project.WebBasedCorrectionEnabled Then
            THIN_CLIENT = True
         End If
      Case Else
         Logging_ExecutionModeString = "BeyondUnknown (" & Project.ScriptExecutionMode & ")"
   End Select

   If THIN_CLIENT Then
      Logging_ExecutionModeString = Logging_ExecutionModeString & " (TC)"
   End If

   Exit Function

   catch:
   Logging_ExecutionModeString = "Unknown Module (Error " & Err.Number & ")"
End Function


'return [classname|subname#linenum]
'  input is the return of WinWrap's CallersLine function: "[macroname|subname#linenum] linetext"
'  refer to the WinWrap documentation for "CallersLine Function" regarding the Depth parameter
Public Function Logging_StackLine(ByVal Caller As String) As String
   On Error GoTo catch

   'the function name (subname) and linenum are followed by a ]
   Dim EndPos As Integer
   EndPos = InStr(Caller, "]")

   'the function name will follow a |
   Dim StartPos As Integer
   StartPos = InStrRev(Caller, "|", EndPos) + 1

   'get the function name
   Dim FunctionAndLine As String
   FunctionAndLine = Mid(Caller, StartPos, EndPos - StartPos)

   'combine with class/folder
   Logging_StackLine = "[" & Logging_SheetClass(Caller) & "|" & FunctionAndLine & "]"

   Exit Function

   catch:
   FunctionAndLine = "Unknown Function (Error " & Err.Number & ")"
End Function


'return the name of the folder or class of the script at the given depth
'  input is the return of WinWrap's CallersLine function: "[macroname|subname#linenum] linetext"
'  refer to the WinWrap documentation for "CallersLine Function" regarding the Depth parameter
Public Function Logging_SheetClass(ByVal Caller As String) As String
   On Error GoTo catch

   'the sheet name (macroname) is followed by a |
   Dim EndPos As Integer
   EndPos = InStr(Caller, "|")

   'the sheet name will follow a \ from Project Builder
   'Project Script: [C:\ProjectFolder\ScriptProject|Document_BeforeProcessXDoc#827] 'Code
   'Other Classes: [C:\1|ValidationForm_ButtonClicked# 18] 'Code
   Dim StartPosPB As Integer
   StartPosPB = InStrRev(Caller, "\", EndPos) + 1

   'the sheet name will follow a * from runtime modules
   '[*ScriptProject|Document_BeforeProcessXDoc#881] 'Code
   Dim StartPosRuntime As Integer
   StartPosRuntime = InStrRev(Caller, "*", EndPos) + 1

   'Use whichever start position is found
   Dim StartPos As Integer
   If StartPosPB > StartPosRuntime Then
      StartPos = StartPosPB
   Else
      StartPos = StartPosRuntime
   End If

   'get the sheet name
   Dim Sheet As String
   Sheet = Mid(Caller, StartPos, EndPos - StartPos)

   'numeric sheet names should be classes or folders
   If IsNumeric(Sheet) Then
      Dim SheetNum As Long
      SheetNum = CLng(Sheet)


      'sheet numbers higher than zero are classes
      If SheetNum > 0 Then
         Dim TheClass As CscClass
         Set TheClass = Project.ClassByID(SheetNum)

         'make sure the class actually exists to prevent an error accessing the name
         If Not TheClass Is Nothing Then
            Logging_SheetClass = TheClass.Name
         Else
            Logging_SheetClass = "Unknown Class (" & SheetNum & ")"
         End If
      Else
         'negative sheet numbers are folders (use absolute value for folder level)
         SheetNum = Abs(SheetNum)
         Dim TheFolder As CscFolderDef
         Set TheFolder = Project.FolderByLevel(SheetNum)

         'make sure the folder actually exists to prevent an error accessing the name
         If Not TheFolder Is Nothing Then
            Logging_SheetClass = TheFolder.Name
         Else
            Logging_SheetClass = "Unknown Folder (" & SheetNum & ")"
         End If
      End If
   ElseIf Sheet = "ScriptProject" Then
      'Project level script has the special designation "ScriptProject"
      Logging_SheetClass = "Project"
   Else
      Logging_SheetClass = "Unknown Class (" & Sheet & ")"
   End If

   Exit Function

   catch:
   Logging_SheetClass = "Unknown Class (Error " & Err.Number & ")"
End Function

'Meant to be called from Batch_Close, this will log routing, rejection, and other details
Public Sub Logging_BatchClose(ByVal pXRootFolder As CASCADELib.CscXFolder, _
   ByVal CloseMode As CASCADELib.CscBatchCloseMode)
   On Error GoTo catch

   Select Case CloseMode
      'routing is evaluated after Final, Suspend, and Error modes
      Case CscBatchCloseMode.CscBatchCloseError
         ErrorLog(Err, "Closing batch in error:" & BATCH_ID & "/" & BATCH_ID_HEX & ", " & _
            BATCH_NAME, Nothing, pXRootFolder, 0, FORCE_ERROR, SUPPRESS_MSGBOX)
         Logging_Routing(pXRootFolder)

         'find any rejected docs
         Dim RejectedMsg As String
         Logging_RejectedDocs(pXRootFolder, RejectedMsg, pXRootFolder.XValues)

         'if there are rejected docs/pages
         If RejectedMsg <> "" Then
            RejectedMsg = "The following have been rejected: " & vbNewLine & RejectedMsg
            ErrorLog(Err, RejectedMsg, Nothing, pXRootFolder, 0, FORCE_ERROR, SUPPRESS_MSGBOX)

            'Potentially take extra action if there is a script error
            '   (set by Logging_RejectedDocs)
            If pXRootFolder.XValues.ItemExists("LOGGING_SCRIPT_ERROR") Then
               'script error action
            End If
         End If

      Case CscBatchCloseMode.CscBatchCloseSuspend
         ScriptLog("Suspending Batch:" & BATCH_ID & "/" & BATCH_ID_HEX & ", " & _
            BATCH_NAME, LOCAL_LOG)
         Logging_Routing(pXRootFolder)

      Case CscBatchCloseMode.CscBatchCloseFinal
         ScriptLog("Batch Close")
         Logging_Routing(pXRootFolder)

      Case CscBatchCloseMode.CscBatchCloseParent
         'Application_InitializeBatch is not called between Child and Parent Batch_Close (SPR00093890)
         '  so initialize logging paths again otherwise Parent logging will go to the Child log
         Logging_InitializeBatch(pXRootFolder)

         'Log that we have "opened" the parent batch
         Logging_BatchOpen(pXRootFolder)
         ScriptLog("Routing complete, closing parent batch.")

      Case CscBatchCloseMode.CscBatchCloseChild
         'Note that if a child batch has been routed to a new batch class,
         '   this will Batch_Close will not fire for the child

         'Log that we have "opened" the child batch
         Logging_BatchOpen(pXRootFolder)

         'See if we can find out which tag this was created with during routing
         Dim i As Integer
         Dim BatchTag As String
         For i = 0 To pXRootFolder.XValues.Count - 1
            If Mid(pXRootFolder.XValues.ItemByIndex(i).Key, 1, _
               Len("KTM_DOCUMENTROUTING_QUEUE_")) = "KTM_DOCUMENTROUTING_QUEUE_" Then
               BatchTag = Mid(pXRootFolder.XValues.ItemByIndex(i).Key, _
                  Len("KTM_DOCUMENTROUTING_QUEUE_") + 1)
               ScriptLog("This batch has been created as a result of routing with " & _
                  "the tag: " & BatchTag)
            End If
         Next
         If BatchTag = "" Then
            ScriptLog("This batch has been created as a result of routing.")
         End If

      Case Else
         ErrorLog(Err, "Unknown Batch Close Type!", Nothing, pXRootFolder, 0, _
            FORCE_ERROR, SUPPRESS_MSGBOX)
   End Select

   Exit Sub

   'if there is an error, log it and try to keep going
   catch:
   ErrorLog(Err, "", Nothing, pXRootFolder, 0, ONLY_ON_ERROR, SUPPRESS_MSGBOX)
   Resume Next
End Sub

'Recursive function to check for rejected docs/pages, called from Logging_BatchClose
Public Sub Logging_RejectedDocs(ByVal XFolder As CscXFolder, ByRef msg As String, _
   ByRef XValues As CscXValues)
   On Error GoTo catch


   Dim i As Integer

   'recurse into folders
   For i = 0 To XFolder.Folders.Count - 1
      Logging_RejectedDocs(XFolder.Folders.ItemByIndex(i), msg, XValues)
   Next

   Dim RejectionNote As String

   'check documents
   Dim XDocInfo As CscXDocInfo
   For i = 0 To XFolder.DocInfos.Count - 1
      Set XDocInfo = XFolder.DocInfos.ItemByIndex(i)

      'check if the doc is rejected
      If XDocInfo.XValues.ItemExists("AC_REJECTED_DOCUMENT") Then
         'identify doc
         msg = msg & Logging_IdentifyFolderDocPage(Nothing, XDocInfo.XDocument)

         'add rejection note if exists
         If XDocInfo.XValues.ItemExists("AC_REJECTED_DOCUMENT_NOTE") Then
            RejectionNote = XDocInfo.XValues.ItemByName("AC_REJECTED_DOCUMENT_NOTE").Value
            msg = msg & ": " & RejectionNote & vbNewLine

            'if the rejection note mentions (S/s)cript,
            '   note for later that there has been a script error
            If InStr(1, RejectionNote, "cript") > 0 Then
               XValues.Set("LOGGING_SCRIPT_ERROR", "True")
            End If
         Else
            msg = msg & vbNewLine
         End If
      End If

      'check pages
      Dim PageIndex As Long
      For PageIndex = 0 To XDocInfo.PageCount - 1
         'check if the page is rejected
         If XDocInfo.XValues.ItemExists("AC_REJECTED_PAGE" & CStr(PageIndex + 1)) Then
            'identify page
            msg = msg & Logging_IdentifyFolderDocPage(Nothing, XDocInfo.XDocument, PageIndex + 1)

            'add rejection note if exists
            If XDocInfo.XValues.ItemExists("AC_REJECTED_PAGE_NOTE" & CStr(PageIndex + 1)) Then
               RejectionNote = XDocInfo.XValues.ItemByName("AC_REJECTED_PAGE_NOTE" & _
                  CStr(PageIndex + 1)).Value
               msg = msg & ": " & RejectionNote & vbNewLine
            Else
               msg = msg & vbNewLine
            End If
         End If
      Next

   Next

   'on error log and exit
   catch:
   ErrorLog(Err, "", Nothing, Nothing, 0, ONLY_ON_ERROR, SUPPRESS_MSGBOX)
End Sub


'Called from Logging_BatchClose to log documents that will be routed
Public Sub Logging_Routing(ByVal pXRootFolder As CscXFolder)
   On Error GoTo catch

   'This will hold the routing information we find:
   '   key=LOGGING_ROUTING_batchtag, value=folders and docs
   Dim RoutingGroups As CscXValues
   Set RoutingGroups = pXRootFolder.XValues

   'Set a flag saying all documents have been routed
   '  Finding an unrouted document will set this to false
   RoutingGroups.Set("LOGGING_ALLROUTED", "True")

   'recursively check folders for routing, adding results to RoutingGroups
   Logging_RoutingFolder(pXRootFolder, RoutingGroups)

   'If all docs are routed this batch will get deleted, so log this to local log
   If pXRootFolder.XValues.ItemByName("LOGGING_ALLROUTED").Value = "True" Then
      ScriptLog("All documents in the batch appear to be routed.  " & _
         "Batch will be deleted.", LOCAL_LOG)
   End If

   'log if the original batch will be routed to a module
   If RoutingGroups.ItemExists("KTM_DOCUMENTROUTING_QUEUE_THISBATCH") Then
      ScriptLog("This original batch will be routed to " & _
         pXRootFolder.XValues.ItemByName("KTM_DOCUMENTROUTING_QUEUE_THISBATCH").Value)
   End If

   'go through the document routing groups and log details
   Dim msg As String
   Dim BatchTag As String
   Dim i As Integer
   For i = 0 To RoutingGroups.Count - 1
      'if the XValue key begins with "LOGGING_ROUTING_"
      If Mid(RoutingGroups.ItemByIndex(i).Key, 1, _
         Len("LOGGING_ROUTING_")) = "LOGGING_ROUTING_" Then

         'the part after "LOGGING_ROUTING_"
         BatchTag = Mid(RoutingGroups.ItemByIndex(i).Key, Len("LOGGING_ROUTING_") + 1)

         msg = msg & "Routing group (" & BatchTag

         'check if it is being routed to a specific queue
         If pXRootFolder.XValues.ItemExists("KTM_DOCUMENTROUTING_QUEUE_" & BatchTag) Then
            msg = msg & ", Queue=" & _
               pXRootFolder.XValues.ItemByName("KTM_DOCUMENTROUTING_QUEUE_" & BatchTag).Value
         End If

         'check if it is being routed with a specific batch name KTM 5.5+
         If pXRootFolder.XValues.ItemExists("KTM_DOCUMENTROUTING_BATCHNAME_" & BatchTag) Then
            msg = msg & ", Batch Name=" & _
               pXRootFolder.XValues.ItemByName("KTM_DOCUMENTROUTING_BATCHNAME_" & BatchTag).Value
         End If

         'check if it is being routed to a new batch class KTM 5.5+
         If pXRootFolder.XValues.ItemExists("KTM_DOCUMENTROUTING_NEWBATCHCLASS_" & BatchTag) Then
            msg = msg & ", Batch Class=" & _
               pXRootFolder.XValues.ItemByName("KTM_DOCUMENTROUTING_NEWBATCHCLASS_" & _
               BatchTag).Value & _
               " (module will be ignored)"
         End If

         msg = msg & "): " & RoutingGroups.ItemByIndex(i).Value & vbNewLine
      End If
   Next

   'if there were any routing groups, msg won't be empty
   If msg <> "" Then
      ScriptLog(msg)
   End If


   'on error log and exit
   catch:
   ErrorLog(Err, "", Nothing, Nothing, 0, ONLY_ON_ERROR, SUPPRESS_MSGBOX)
End Sub


'recursively check folders for routed documents (or first level routed folders),
'   adding results to RoutingGroups
Public Sub Logging_RoutingFolder(ByVal XFolder As CscXFolder, ByRef RoutingGroups As CscXValues)
   On Error GoTo catch

   'only 1st level folders can be routed (but any documents can be routed)
   Dim IsFirstLevelFolder As Boolean
   IsFirstLevelFolder = False

   If XFolder.IsRootFolder = False Then 'not the root
      If XFolder.ParentFolder.IsRootFolder = True Then 'parent is the root
         IsFirstLevelFolder = True
      End If
   End If

   Dim BatchTag As String

   'check if this folder is being routed
   If IsFirstLevelFolder And XFolder.XValues.ItemExists("KTM_DOCUMENTROUTING") Then
      BatchTag = XFolder.XValues.ItemByName("KTM_DOCUMENTROUTING").Value

      Dim FolderName As String
      FolderName = Logging_IdentifyFolderDocPage(XFolder)

      'check if we've already added this group
      If RoutingGroups.ItemExists("LOGGING_ROUTING_" & BatchTag) Then
         'add this folder
         RoutingGroups.Set("LOGGING_ROUTING_" & BatchTag, _
            RoutingGroups.ItemByName("LOGGING_ROUTING_" & BatchTag).Value & "," & FolderName)
      Else
         'create it and add this document
         RoutingGroups.Set("LOGGING_ROUTING_" & BatchTag, FolderName)
      End If

      'if the folder is being routed, it will route the contents,
      '  and if routing instructions were set on these contents, they will be ignored
   Else
      'if the folder is not being routed, check if its subfolders
      Dim SubFolder As CscXFolder
      Dim i As Integer
      For i = 0 To XFolder.Folders.Count - 1
         Set SubFolder = XFolder.Folders.ItemByIndex(i)
         Logging_RoutingFolder(SubFolder, RoutingGroups)
      Next

      Dim oXDocInfo As CscXDocInfo
      Dim DocName As String

      'check for routed docs in this folder
      For i = 0 To XFolder.DocInfos.Count - 1
         Set oXDocInfo = XFolder.DocInfos.ItemByIndex(i)
         DocName = Logging_IdentifyFolderDocPage(Nothing, oXDocInfo.XDocument)

         'check if this document is being routed
         If oXDocInfo.XValues.ItemExists("KTM_DOCUMENTROUTING") Then
            BatchTag = oXDocInfo.XValues.ItemByName("KTM_DOCUMENTROUTING").Value

            'check if we've already added this group
            If RoutingGroups.ItemExists("LOGGING_ROUTING_" & BatchTag) Then
               'add this document
               RoutingGroups.Set("LOGGING_ROUTING_" & BatchTag, _
                  RoutingGroups.ItemByName("LOGGING_ROUTING_" & BatchTag).Value & "," & DocName)
            Else
               'create it and add this document
               RoutingGroups.Set("LOGGING_ROUTING_" & BatchTag, DocName)
            End If
         Else
            'If a document is not being routed, we know they are not all routed
            RoutingGroups.Set("LOGGING_ALLROUTED", "False")
         End If
      Next
   End If


   'on error log and exit
   catch:
   ErrorLog(Err, "", Nothing, Nothing, 0, ONLY_ON_ERROR, SUPPRESS_MSGBOX)
End Sub


'Identifies structure and files from Folder/Doc/Page.
Public Function Logging_IdentifyFolderDocPage(Optional ByVal XFolder As CscXFolder = Nothing, _
   Optional ByVal XDoc As CscXDocument = Nothing, _
   Optional ByVal PageNum As Integer = 0) As String
   'valid parameter combinations:
   'only folder
   'only doc, implies folder
   'doc and page number (page object doesn't link to parent doc)

   On Error GoTo catch

   'if doc was provided, use that to set folder
   If Not XDoc Is Nothing Then
      Set XFolder = XDoc.ParentFolder
   Else
      'if doc was not provided, make sure we have a folder, or exit
      If XFolder Is Nothing Then
         Exit Function
      Else
         'also exit if we only have root because we are omitting root folder info
         If XFolder.IsRootFolder Then
            Exit Function
         End If
      End If
   End If

   'Structure will show info like F#\F#\D#\P#
   Dim DocStructure As String

   'Files will show info like (#.xdc\#.tif(#))
   '  (#) is the page number in that document
   '  All folders are in Folder.xfd, so no need to add that
   Dim Files As String


   '2013-08-08 - SMK - XDoc.ParentFolder is not always set in KTM 6.0 Project Builder
   '  In that case, just skip this section
   If Not XFolder Is Nothing Then
      'get the folder structure (other than root folder since it is always there)
      Do While Not XFolder.IsRootFolder

         Set XFolder = XFolder.ParentFolder
         DocStructure = "F" & XFolder.IndexInFolder + 1 & "\" & DocStructure
         Files = Mid(XFolder.FileName, InStrRev(XFolder.FileName, "\")) & Files
      Loop
   End If

   'get the document
   If Not XDoc Is Nothing Then
      DocStructure = DocStructure & "D" & XDoc.IndexInFolder + 1
      Files = Files & Mid(XDoc.FileName, InStrRev(XDoc.FileName, "\") + 1)
   End If

   'get the page
   If PageNum > 0 And Not XDoc Is Nothing Then
      DocStructure = DocStructure & "\P" & PageNum

      Dim Page As CscXDocPage
      Set Page = XDoc.Pages.ItemByIndex(PageNum - 1)
      'make sure the page exists
      If Not Page Is Nothing Then
         Files = Files & Mid(Page.SourceFileName, InStrRev(Page.SourceFileName, "\"))
      End If
   End If

   Logging_IdentifyFolderDocPage = DocStructure & " (" & Files & ")"


   Exit Function

   catch:
   Logging_IdentifyFolderDocPage = "Unknown Folder/Doc/Page"
End Function


'Wrapper and drop-in replacement for MsgBox.
Public Function MsgBoxLog(ByVal Message As String, _
   Optional ByVal MsgType As VbMsgBoxStyle, _
   Optional ByVal Title As String, _
   Optional ByVal pXDoc As CscXDocument=Nothing, _
   Optional ByVal pXFolder As CscXFolder=Nothing, _
   Optional PageNum As Integer=0) As VbMsgBoxResult

   On Error GoTo catch

   'Figure out what kind of MsgBox style this is (one from each group)
   Dim TypeString As String

   'Buttons
   If CInt(MsgType And vbOkOnly) = vbOkOnly Then
      TypeString = "vbOkOnly, "
   ElseIf CInt(MsgType And vbOkCancel) = vbOkCancel Then
      TypeString = "vbOkCancel, "
   ElseIf CInt(MsgType And vbAbortRetryIgnore) = vbAbortRetryIgnore Then
      TypeString = "vbAbortRetryIgnore, "
   ElseIf CInt(MsgType And vbYesNoCancel) = vbYesNoCancel Then
      TypeString = "vbYesNoCancel, "
   ElseIf CInt(MsgType And vbYesNo) = vbYesNo Then
      TypeString = "vbYesNo, "
   ElseIf CInt(MsgType And vbRetryCancel) = vbRetryCancel Then
      TypeString = "vbRetryCancel, "
   End If

   'Icon
   If CInt(MsgType And vbCritical) = vbCritical Then
      TypeString = TypeString & "vbCritical, "
   ElseIf CInt(MsgType And vbQuestion) = vbQuestion Then
      TypeString = TypeString & "vbQuestion, "
   ElseIf CInt(MsgType And vbExclamation) = vbExclamation Then
      TypeString = TypeString & "vbExclamation, "
   ElseIf CInt(MsgType And vbInformation) = vbInformation Then
      TypeString = TypeString & "vbInformation, "
   End If

   'Default
   If CInt(MsgType And vbDefaultButton1) = vbDefaultButton1 Then
      TypeString = TypeString & "vbDefaultButton1, "
   ElseIf CInt(MsgType And vbDefaultButton2) = vbDefaultButton2 Then
      TypeString = TypeString & "vbDefaultButton2, "
   ElseIf CInt(MsgType And vbDefaultButton3) = vbDefaultButton3 Then
      TypeString = TypeString & "vbDefaultButton3, "
   End If

   'Default
   If CInt(MsgType And vbApplicationModal) = vbApplicationModal Then
      TypeString = TypeString & "vbApplicationModal"
   ElseIf CInt(MsgType And vbSystemModal) = vbSystemModal Then
      TypeString = TypeString & "vbSystemModal"
   ElseIf CInt(MsgType And vbMsgBoxSetForeground) = vbMsgBoxSetForeground Then
      TypeString = TypeString & "vbMsgBoxSetForeground"
   End If


   'log an error as a warning if this is used during server
   If Project.ScriptExecutionMode = CscScriptExecutionMode.CscScriptModeServer Then
      ErrorLog(Err, "Skipping a MsgBox and forcing an OK result because " & _
         "it is running during Server. Message: """ & Message & """ (" & TypeString & _
         ")", pXDoc, pXFolder, PageNum, FORCE_ERROR, SUPPRESS_MSGBOX)
      MsgBoxLog = vbOK
      Exit Function
   End If

   'MsgBox also cannot be used from a Thin Client
   If THIN_CLIENT Then
      ErrorLog(Err, "Skipping a MsgBox and forcing an OK result because " & _
         "it is running during Thin Client. Message: """ & Message & """ (" & TypeString & _
         ")", pXDoc, pXFolder, PageNum, FORCE_ERROR, SUPPRESS_MSGBOX)
      MsgBoxLog = vbOK
      Exit Function
   End If

   'show the message and grab the result
   Dim Result As VbMsgBoxResult
   Result = MsgBox(Message, MsgType, Title)

   'Find out what the user clicked
   Dim ResultString As String
   Select Case Result
      Case vbOK
         ResultString = "OK"
      Case vbCancel
         ResultString = "Cancel"
      Case vbAbort
         ResultString = "Abort"
      Case vbRetry
         ResultString = "Retry"
      Case vbIgnore
         ResultString = "Ignore"
      Case vbYes
         ResultString = "Yes"
      Case vbNo
         ResultString = "No"
      Case Else
         ResultString = "Unknown"
   End Select

   'log the details
   ScriptLog("MsgBox: User clicked " & ResultString & " for message """ & Message & """ (" & _
      TypeString & ")", BATCH_LOG, pXDoc, pXFolder, PageNum, IGNORE_CURRENT_FUNCTION)

   'return the result just like a normal MsgBox
   MsgBoxLog = Result

   Exit Function

   'if there is an error, log it and try to keep going
   catch:
   ErrorLog(Err, "", pXDoc, pXFolder, 0, ONLY_ON_ERROR, SUPPRESS_MSGBOX)
   Resume Next
End Function
'========  END   LOGGING CODE ========

Private Sub Application_InitializeBatch(ByVal pXRootFolder As CASCADELib.CscXFolder)
   Logging_InitializeBatch(pXRootFolder)
End Sub



Public Sub Logging_Classification(pXDoc As CscXDocument, Optional Popup As Boolean=False)
   Dim msg As String
   msg=Logging_StackLine(CallersLine(0)) & vbNewLine
   msg=msg & Logging_ClassifcationResult("Classification - Document " & pXDoc.IndexInFolder + 1 & ": ", pXDoc.ClassificationResult,"Document")
   Dim i As Integer
   For i=0 To pXDoc.CDoc.Pages.Count-1
      If pXDoc.CDoc.Pages(i).SplitPage Then
         msg=msg & "Start of New Document (SplitPage = TRUE)" & vbNewLine
      End If

      msg = msg & "Page Classification - Page " & i & ": " & pXDoc.CDoc.Pages(i).PageClassificationResult & " = " & pXDoc.CDoc.Pages(i).PageClassificationConfidence & vbNewLine
      msg=msg & Logging_ClassifcationResult("Layout Classification - Page " & i & ": ",pXDoc.CDoc.Pages(i).ClassificationResult,"Layout")
      If pXDoc.Representations.Count > 0 Then
         msg=msg & Logging_ClassifcationResult("Content Classification - Page " & i & ": ",pXDoc.Pages(i).ClassificationResult,"Content")
      End If
      msg=msg & Logging_ClassifcationResult("TDS Classification - Page " & i & ": ",pXDoc.CDoc.Pages(i).TDSAlternativeResults,"TDSPage")
   Next

   If Popup=False Then
      ScriptLog(msg,BATCH_LOG,pXDoc,Nothing,0,IGNORE_CURRENT_FUNCTION)
   Else
      MsgBoxLog(msg,vbInformation,"Classification Result",pXDoc,Nothing,0)
   End If
End Sub



'returns information about a specific classification confidence. * means it passes the applicable threshold
Public Function Logging_ConfidenceLine(Confidence As Single, ktmClass As CscClass, ResultType As CscResultType, IsTDSResult As Boolean, ResultFrom As String) As String
   If Confidence<0.001 Then Exit Function

   Dim msg As String
   msg=Format(Confidence,"Percent") & " - " & ktmClass.Name & "(" & Logging_ResultTypeString(ResultType) & ")"

   'threshold to use, class or project
   Dim Threshold As Single

   Select Case ResultFrom
      Case "Document"
         If IsTDSResult Then
            'TDS separation threshold only decides if the ClassificationUnconfident flag will be set to make it unconfident in Doc Review
            If ktmClass.UseProjectThresholdsOnly Then
               'Currently the TDS threshold reuses the project Content Confidence value
               Threshold=Project.MinContentConfidence
            Else
               Threshold=ktmClass.MinClsConfidenceTDS

               'note the class specific threshold
               msg=msg & " (ClsThreshold:" & Format(Threshold,"Percent") & ")"
            End If

            'note if it is above the threshold
            If Confidence>Threshold Then
               msg=msg & "*"
            End If
         Else
            'For non TDS, I believe the document level just contains a copy of the results from whicheer classifier the result was from
         End If
      Case "Layout"
         If IsTDSResult Then
            'TDS classifiers do not obey class thresholds, only project
            If Confidence>Project.MinLayoutConfidence Then
               msg=msg & "*"
            End If
         Else
            If ktmClass.UseProjectThresholdsOnly Then
               Threshold=Project.MinLayoutConfidence
            Else
               Threshold=ktmClass.MinClsConfidenceLayout

               'note the class specific threshold
               msg=msg & " (ClsThreshold:" & Format(Threshold,"Percent") & ")"
            End If

            'note if it is above the threshold
            If Confidence>Threshold Then
               msg=msg & "*"
            End If
         End If
      Case "Content"
         If IsTDSResult Then
            'Content results from TDS only show up if they are past the hardcoded 10% threshold
            msg=msg & "*"
         Else
            If ktmClass.UseProjectThresholdsOnly Then
               Threshold=Project.MinContentConfidence
            Else
               Threshold=ktmClass.MinClsConfidenceContent

               'note the class specific threshold
               msg=msg & " (ClsThreshold:" & Format(Threshold,"Percent") & ")"
            End If

            'note if it is above the threshold
            If Confidence>Threshold Then
               msg=msg & "*"
            End If
         End If
      Case Else
         'nothing
   End Select

   Logging_ConfidenceLine=msg
End Function




'Returns string from CscResultType Enum
Public Function Logging_ResultTypeString(rt As CscResultType) As String
   On Error Resume Next

   Dim strRT As String
   strRT=""

   'Because the types can be combined, test by using logical AND on rt and the type
   If CInt(rt And CscResultTypeFinalResult) = CscResultTypeFinalResult Then
         strRT=strRT & "FinalResult, " ' (" & CscResultTypeFinalResult & "), "
   End If
   If CInt(rt And CscResultTypeWeighted) = CscResultTypeWeighted Then
         strRT=strRT & "Weighted, " ' (" & CscResultTypeWeighted & "), "
   End If
   If CInt(rt And CscResultTypeLocalNo) = CscResultTypeLocalNo Then
         strRT=strRT & "LocalNo, " ' (" & CscResultTypeLocalNo & "), "
   End If
   If CInt(rt And CscResultTypeLocalYes) = CscResultTypeLocalYes Then
         strRT=strRT & "LocalYes, " ' (" & CscResultTypeLocalYes & "), "
   End If
   If CInt(rt And CscResultTypeLocalValid) = CscResultTypeLocalValid Then
         strRT=strRT & "LocalValid, " ' (" & CscResultTypeLocalValid & "), "
   End If
   If CInt(rt And CscResultTypeRedirected) = CscResultTypeRedirected Then
         strRT=strRT & "Redirected, " ' (" & CscResultTypeRedirected & "), "
   End If
   If CInt(rt And CscResultTypeDefault) = CscResultTypeDefault Then
         strRT=strRT & "Default, " ' (" & CscResultTypeDefault & "), "
   End If
   If CInt(rt And CscResultTypePropagatedNo) = CscResultTypePropagatedNo Then
         strRT=strRT & "PropagatedNo, " ' (" & CscResultTypePropagatedNo & "), "
   End If
   If CInt(rt And CscResultTypeSubtree) = CscResultTypeSubtree Then
         strRT=strRT & "Subtree, " ' (" & CscResultTypeSubtree & "), "
   End If
   If CInt(rt And CscResultTypeParentRepChilds) = CscResultTypeParentRepChilds Then
         strRT=strRT & "ParentRepChilds, " ' (" & CscResultTypeParentRepChilds & "), "
   End If
   If CInt(rt And CscResultTypeChildOverrulesParent) = CscResultTypeChildOverrulesParent Then
         strRT=strRT & "ChildOverrulesParent, " ' (" & CscResultTypeChildOverrulesParent & "), "
   End If
   If CInt(rt And CscResultTypeReclassification) = CscResultTypeReclassification Then
         strRT=strRT & "Reclassification, " ' (" & CscResultTypeReclassification & "), "
   End If
   If CInt(rt And CscResultTypeFirstPage) = CscResultTypeFirstPage Then
         strRT=strRT & "FirstPage, " ' (" &CscResultTypeFirstPage  & "), "
   End If
   If CInt(rt And CscResultTypeMiddlePage) = CscResultTypeMiddlePage Then
         strRT=strRT & "MiddlePage, " ' (" & CscResultTypeMiddlePage & "), "
   End If
   If CInt(rt And CscResultTypeLastPage) = CscResultTypeLastPage Then
         strRT=strRT & "LastPage, " ' (" & CscResultTypeLastPage & "), "
   End If

   'remove the last comma
   strRT=Mid(strRT,1,Len(strRT)-2)

   'check if there is still a comma, which means more than one type combined
   Dim CommaPos As Integer
   CommaPos=InStr(1,strRT,",")

   'if there are no types or multiple types, also output the type number
   If CommaPos>0 Or strRT="" Then
      strRT=strRT & "(" & rt & ")"
   End If

   Logging_ResultTypeString=strRT
End Function





Public Function Logging_ClassifcationResult(Header As String, ByVal CR As CscResult, ResultsFrom As String) As String
   'On Error Resume Next
   If Not CR Is Nothing Then

      Dim msg As String
      msg=""

      'should this result be considered in context of TDS
      Dim IsTDSResult As Boolean

      'draw some conclusions from the reptype
      Select Case CR.RepType
         Case "Fres"
            '"Full", observed at Document level
            msg=msg & "Document (Fres) results:" & vbNewLine

            'if TDS (ADS) is enabled in the project consider this a tds result
            If Project.ADSEnabled Then
               IsTDSResult=True
            Else
               IsTDSResult=False
            End If
         Case "pres"
            'Page level - layout/content classifiers
            IsTDSResult=False
            msg=msg & ResultsFrom & " classifier results:" & vbNewLine
         Case "Sres"
            'Means "sparse" but observed in TDS... Separation - TDS
            IsTDSResult=True
            msg=msg & ResultsFrom & " TDS results:" & vbNewLine
         Case Else
            msg=msg & "Unknown RepType: " & CR.RepType & " from " & ResultsFrom & vbNewLine
      End Select


      Dim i As Integer
      Dim CurrentConfidence As Single
      Dim CurrentClassId As Long
      Dim CurrentResultType As CscResultType
      Dim TDSClassesLogged As String ' ,ClassID,
      Dim ktmClass As CscClass

      For i=0 To CR.FinalClassCount-1
         'Get current ClassId and Confidence
         CurrentClassId = CR.FinalClassId(i)
         CurrentConfidence = CR.Confidence(CurrentClassId) 'is this the right conf?

         If CurrentConfidence>0.001 Then
            'This returns an arbitary type when a TDS result has different page type results from the same class
            CurrentResultType=CR.ResultType(CurrentClassId)

            Set ktmClass=Project.ClassByID(CurrentClassId)
            If Not ktmClass Is Nothing Then
               'Check TDS page types
               If CurrentResultType=CscResultTypeFirstPage Or _
                  CurrentResultType=CscResultTypeMiddlePage Or _
                  CurrentResultType=CscResultTypeLastPage Then

                  'Since we log all types at once, don't log them again when we see the next type
                  'If InStr(1,TDSClassesLogged,"," & CurrentClassId & ",")=0 Then
                     Dim ConfFirst As Single
                     Dim ConfMiddle As Single
                     Dim ConfLast As Single

                     'use GetPageLevelConfidence for each type
                     ConfFirst=CR.GetPageLevelConfidence(CurrentClassId,CscResultTypeFirstPage)
                     ConfMiddle=CR.GetPageLevelConfidence(CurrentClassId,CscResultTypeMiddlePage)
                     ConfLast=CR.GetPageLevelConfidence(CurrentClassId,CscResultTypeLastPage)

                     'log each type that has any confidence
                     If ConfFirst=CurrentConfidence Then msg=msg & Header & "(Final) " & Logging_ConfidenceLine(ConfFirst,ktmClass,CscResultTypeFirstPage,IsTDSResult,ResultsFrom) & vbNewLine
                     If ConfMiddle=CurrentConfidence Then msg=msg & Header & "(Final) " & Logging_ConfidenceLine(ConfMiddle,ktmClass,CscResultTypeMiddlePage,IsTDSResult,ResultsFrom) & vbNewLine
                     If ConfLast=CurrentConfidence Then msg=msg & Header & "(Final) " & Logging_ConfidenceLine(ConfLast,ktmClass,CscResultTypeLastPage,IsTDSResult,ResultsFrom) & vbNewLine

                     'Remember that we did this class, so we don't log them again
                     'TDSClassesLogged=TDSClassesLogged & "," & CurrentClassId & ","
                  'End If
               Else
                  'Other types are simpler
                  msg=msg & Header & "(Final) " & Logging_ConfidenceLine(CurrentConfidence,ktmClass,CurrentResultType,IsTDSResult,ResultsFrom) & vbNewLine
               End If
            Else
               msg=msg & Header  & "(Final) " & "WARNING id " & CurrentClassId & " with confidence " & CurrentConfidence & " did not return a valid class from Project.ClassByID!" & vbNewLine
            End If
         End If

      Next


      TDSClassesLogged=""

      'Check all confidences in the CR
      For i = 0 To CR.NumberOfConfidences-1
         'Get current ClassId and Confidence
         CurrentConfidence = CR.BestConfidence(i)
         CurrentClassId = CR.BestClassId(i)

         If CurrentConfidence>0.001 Then

            'This returns an arbitary type when a TDS result has different page type results from the same class
            CurrentResultType=CR.ResultType(CurrentClassId)

            Set ktmClass=Project.ClassByID(CurrentClassId)
            If Not ktmClass Is Nothing Then
               'Check TDS page types
               If CurrentResultType=CscResultTypeFirstPage Or _
                  CurrentResultType=CscResultTypeMiddlePage Or _
                  CurrentResultType=CscResultTypeLastPage Then

                  'Since we log all types at once, don't log them again when we see the next type
                  'If InStr(1,TDSClassesLogged,"," & CurrentClassId & ",")=0 Then
                     'Dim ConfFirst As Single
                     'Dim ConfMiddle As Single
                     'Dim ConfLast As Single

                     'use GetPageLevelConfidence for each type
                     ConfFirst=CR.GetPageLevelConfidence(CurrentClassId,CscResultTypeFirstPage)
                     ConfMiddle=CR.GetPageLevelConfidence(CurrentClassId,CscResultTypeMiddlePage)
                     ConfLast=CR.GetPageLevelConfidence(CurrentClassId,CscResultTypeLastPage)

                     'log each type that has any confidence
                     If ConfFirst=CurrentConfidence Then msg=msg & Header & Logging_ConfidenceLine(ConfFirst,ktmClass,CscResultTypeFirstPage,IsTDSResult,ResultsFrom) & vbNewLine
                     If ConfMiddle=CurrentConfidence Then msg=msg & Header & Logging_ConfidenceLine(ConfMiddle,ktmClass,CscResultTypeMiddlePage,IsTDSResult,ResultsFrom) & vbNewLine
                     If ConfLast=CurrentConfidence Then msg=msg & Header & Logging_ConfidenceLine(ConfLast,ktmClass,CscResultTypeLastPage,IsTDSResult,ResultsFrom) & vbNewLine

                     'Remember that we did this class, so we don't log them again
                     'TDSClassesLogged=TDSClassesLogged & "," & CurrentClassId & ","
                  'End If
               Else
                  'Other types are simpler
                  msg=msg & Header & Logging_ConfidenceLine(CurrentConfidence,ktmClass,CurrentResultType,IsTDSResult,ResultsFrom) & vbNewLine
               End If
            Else
               msg=msg & Header & "WARNING id " & CurrentClassId & " with confidence " & CurrentConfidence & " did not return a valid class from Project.ClassByID!" & vbNewLine
            End If
         End If
      Next i

      If CR.NumberOfConfidences>0 Then
         Select Case ResultsFrom
            Case "Document"
               If IsTDSResult Then
                  msg=msg & "*indicates result meets TDS project threshold of " & Format(Project.MinContentConfidence,"Percent") & _
                     " or class specific threshold (Confident in DR)" & vbNewLine
               End If
            Case "Layout"
               msg=msg & "*indicates result meets layout project threshold of " & Format(Project.MinLayoutConfidence,"Percent") & _
                  " or class specific threshold" & vbNewLine
               msg=msg & "Layout results also require a difference of " & Format(Project.MinLayoutDistance,"Percent") & vbNewLine
            Case "Content"
               If IsTDSResult Then
                  'TDS content uses hardcoded threshold and no difference
                  msg=msg & "*Content results from TDS only show up if they meet a hardcoded 10% minimum confidence." & vbNewLine
               Else
                  msg=msg & "*indicates result meets content project threshold of " & Format(Project.MinContentConfidence,"Percent") & _
                     " or class specific threshold" & vbNewLine
                  msg=msg & "Content results also require a difference of " & Format(Project.MinContentDistance,"Percent") & vbNewLine
               End If
            Case Else
               'nothing
         End Select
      Else
         msg=msg & Header & "There are zero confidences in this ClassificationResult" & vbNewLine
      End If
   Else
      If ResultsFrom="Content" Then
         msg=msg & "Content ClassificationResult is nothing, which generally means Layout results produced valid results" & vbNewLine
      Else
         msg=msg & Header & "ClassificationResult is nothing" & vbNewLine
      End If
   End If

   Logging_ClassifcationResult=msg & vbNewLine
End Function





Public Sub Separation_MergeDocsAndSeparate(pXRootFolder As CscXFolder)
   Separation_MergeAllDocs(pXRootFolder)
   Project.SeparatePages(pXRootFolder.DocInfos(0).XDocument)
   Separation_SplitBatchAsMarked(pXRootFolder)
   Separation_ClassifyAllDocs(pXRootFolder)
End Sub

Public Sub Separation_ClassifyAllDocs(pXRootFolder As CscXFolder)
   Dim DocIndex As Long, pXDoc As CscXDocument, Doc As CscXDocInfo
   For DocIndex = 0 To pXRootFolder.DocInfos.Count-1
      Set Doc=pXRootFolder.DocInfos(DocIndex)
      Set pXDoc=Doc.XDocument
      Project.ClassifyXDoc(pXDoc)
   Next
End Sub


Public Sub Separation_MergeAllDocs(pXRootFolder As CscXFolder)
   Dim DocIndex As Integer, CurDoc As CscXDocInfo, PrevDoc As CscXDocInfo

   For DocIndex=pXRootFolder.DocInfos.Count-1 To 1 Step -1 'from last doc to second
      Set CurDoc=pXRootFolder.DocInfos(DocIndex)
      Set PrevDoc=pXRootFolder.DocInfos(DocIndex-1)

      ' skip docs with readonly structure (pdf)
      If Not (CurDoc.IsPageStructureReadOnly Or PrevDoc.IsPageStructureReadOnly) Then
         Batch.MergeDocuments(PrevDoc,CurDoc)
      Else
         Debug.Print("Not able to merge because document has readonly structure (PDF, etc): Document Index " & IIf(CurDoc.IsPageStructureReadOnly,CurDoc.IndexInFolder,PrevDoc.IndexInFolder))
      End If
   Next

   'While pXRootFolder.DocInfos.Count>1
   '   Batch.MergeDocuments(pXRootFolder.DocInfos(0),pXRootFolder.DocInfos(1))
   'Wend
End Sub


Public Sub Separation_SplitBatchAsMarked(pXRootFolder As CscXFolder)
   Dim DocIndex As Long, pXDoc As CscXDocument, Doc As CscXDocInfo
   For DocIndex = pXRootFolder.DocInfos.Count-1 To 0 Step - 1
      Set Doc=pXRootFolder.DocInfos(DocIndex)

      If Not Doc.IsPageStructureReadOnly Then
         Set pXDoc=Doc.XDocument

         Dim PageIndex As Long
         For PageIndex=pXDoc.CDoc.Pages.Count-1 To 1 Step -1 'last page to second page (first can't be "split")
            If pXDoc.CDoc.Pages(PageIndex).SplitPage Then
               Batch.SplitDocumentAfterPage(Doc,PageIndex-1)
            End If
         Next
         Doc.UnloadDocument()
      Else
         Debug.Print("Not able to split because document has readonly structure (PDF, etc): Document Index " & DocIndex)
      End If
   Next
End Sub



Private Declare Function GetModuleFileNameEx Lib "psapi" Alias "GetModuleFileNameExA" ( ByVal hProcess As Long, ByVal hModule As Long, ByVal FileName As String, ByVal nSize As Long) As Long

Public Function IsDesignMode() As Boolean
   ' When testing Runtime Script Event (lightning bolt, Batch_ and Application_ events), ScriptExecutionMode is NOT set to Design
   ' So this function is required for code in these functions that needs to know the difference between runtime and testing in ProjectBuilder
   Dim FileName As String
   FileName = Space$(256)
   GetModuleFileNameEx(-1,0, FileName, 256)
   FileName = Left$(FileName, InStr(1, FileName, ChrW$(0)))
   Return InStr(1, FileName, "ProjectBuilder") ' This works for old PB (KTM 5.5) as well as new PB (KTM 6.0+, KTA)
End Function


Public Function NewScriptFunction(m As Match) As ScriptFunction
   ' Creates a new ScriptFunction from regex match, implementation tied to regex in ParseScript()

   With NewScriptFunction
      .StartIndex=m.FirstIndex
      .EndIndex=m.FirstIndex + m.Length
      .IsSub=(LCase(m.SubMatches(0))="sub")
      .Name=m.SubMatches(1)
      ' provide the string of params to be parsed out into a dictionary of param types
      .Params=NewParams(m.SubMatches(2))
      .returntype=m.SubMatches(3)
      .Content=m
   End With
End Function

Public Function NewParams(ParamString As String) As Param()
   ' Creates array of Param from a parameter string (everything between parens)

   Dim r As New RegExp, Matches As MatchCollection
   r.Global=True:   r.Multiline=True:   r.IgnoreCase=True
   r.Pattern = "(Optional *?)?(ByVal|ByRef)? *(\w+?)(\( *\))?(?: As ([\w\.]+?)) *(?:= *(.*?))?(?:,|$)"
   Set Matches = r.Execute(ParamString)

   If Matches.Count>0 Then
      Dim MatchIndex As Integer, p As Param, Params() As Param
      ReDim Params(Matches.Count-1)
      For MatchIndex=0 To Matches.Count-1
         p = NewParam(Matches(MatchIndex))
         Params(MatchIndex)=p
      Next
   End If
   Return Params
End Function

Public Function NewParam(m As Match) As Param
   ' Creates a Param from regex match, implementation tied to regex in NewParams()

   With NewParam
      .OptionalParam=Len(Trim(m.SubMatches(0)))>0
      .Name=m.SubMatches(2)
      .Array=Len(Trim(m.SubMatches(3)))>0
      .ParamType=m.SubMatches(4)
      .DefaultValue=m.SubMatches(5)
   End With
End Function



Public Function ParseScript(ByVal Script As String, Optional ByVal ClassName As String = "") As ScriptFunction()
   ' Parse script into array of ScriptFunction

   ' Return from cache if possible
   Static Cache As New Dictionary
   If Cache.Exists(ClassName) Then
      Return Cache.Item(ClassName)
   End If

   Dim r As New RegExp, Matches As MatchCollection
   r.Global=True:   r.Multiline=True:   r.IgnoreCase=True
   'vbscript regexp does not support singleline mode (. matches \n) and no support for named capturing groups
   r.Pattern = "^(?:Public |Private )?(Sub|Function) (.*?)\((.*?)\)\s*(?: As (.+?))?$((?:.|\n)*?)End \1"
   Set Matches = r.Execute(Script)
   Dim SFs() As ScriptFunction

   If Matches.Count=0 Then Return SFs
   ReDim SFs(Matches.Count-1)

   Dim MatchIndex As Long, sf As ScriptFunction
   For MatchIndex=0 To Matches.Count-1
      sf=NewScriptFunction(Matches(MatchIndex))
      sf.ClassName=ClassName
      SFs(MatchIndex)=sf
   Next

   ' Add to cache
   Cache.Add(ClassName, SFs)

   Return SFs
End Function


Public Function DevMenu_FilteredFunctions(InputFunctions() As ScriptFunction, ByRef Filtered() As ScriptFunction, Optional IncludeFolder As Boolean=True, Optional IncludeDoc As Boolean=True, Optional IncludeOther As Boolean=True, Optional ContainsStr As String="") As String()
   ' Out variable Filtered provides a filtered array of ScriptFunctions, function returns array of strings corresponding to the function names to use in a menu/dropdown

   Dim Choices() As String, i As Long, f As ScriptFunction, FilterResults As Long, ValidParams As Boolean
   ReDim Choices(0) : ReDim Filtered(0)

   For i=LBound(InputFunctions) To UBound(InputFunctions)
      f=InputFunctions(i)

      ' Match name filter
      If ContainsStr = "" Or InStr(1,f.Name,ContainsStr)>0 Then
         ValidParams=False

         If UBound(f.Params)=-1 OrElse f.Params(0).OptionalParam Then
            If IncludeOther Then ValidParams=True ' No params or all optional
         End If
         If UBound(f.Params)>-1 Then
            Dim p As Param
            p=f.Params(0)
            ' Check for first param of XDoc or XFolder and if that type is included in filter
            If (Replace(LCase(p.ParamType),"cascadelib.","")="cscxfolder" And IncludeFolder) Or _
               (Replace(LCase(p.ParamType),"cascadelib.","")="cscxdocument" And IncludeDoc) Then

               ' Either a single param, or everything after the first is optional
               If UBound(f.Params)=0 OrElse f.Params(1).OptionalParam Then
                  ValidParams=True
               End If
            End If
         End If

         If ValidParams Then
            ReDim Preserve Choices(FilterResults)
            ReDim Preserve Filtered(FilterResults)

            ' Function passes filter so add to filtered functions and menu list
            Choices(FilterResults)=InputFunctions(i).Name
            Filtered(FilterResults)=f
            FilterResults=FilterResults+1
         End If
      End If
   Next

   Return Choices
End Function





Public Sub DevMenu_Dialog(Optional pXFolder As CscXFolder=Nothing, Optional pXDoc As CscXDocument=Nothing)
   ' Show a menu that allows testing functions from the Project script with the provided Folder or Doc

   If Not IsDesignMode() Then Exit Sub
   Debug.Clear

   Dim AllFunc() As ScriptFunction, FilteredFunc() As ScriptFunction
   AllFunc=ParseScript(Project.ScriptCode,"Project") ' TODO: choice for local class script instead of project

   Dim FuncNames() As String
   FuncNames=DevMenu_FilteredFunctions(AllFunc,FilteredFunc)

   ' collect function name prefixes before an underscore to use as preset filter groups
   Dim Prefixes(1000) As String, CurPf As String, Pre As New Dictionary, sf As ScriptFunction
   Prefixes(Pre.Count)="Show All" : Pre.Add("Show All","Show All")
   Prefixes(Pre.Count)="Custom Filter:" : Pre.Add("Custom Filter:","Custom Filter:")
   For Each sf In FilteredFunc
      If UBound(Split(sf.Name,"_"))>0 Then
         CurPf=Split(sf.Name,"_")(0)
         If Not Pre.Exists(CurPf) Then
            Prefixes(Pre.Count)=CurPf
            ' dictionary is only used to check whether we've added something already, contents not used
            Pre.Add(CurPf,CurPf)
         End If
      End If
   Next
   On Error Resume Next
   ' Keep an exported copy of the project script updated.
   ' Called via eval and ignoring errors so this will work if it is in the project, and no error if it is not.
   Eval("Dev_ExportScriptAndLocators()")
   ' Get the name of the calling function
   On Error GoTo 0

   Begin Dialog DevDialog 640,399,"Transformation Script Development Menu",.DevMenu_DialogFunc ' %GRID:10,7,1,1
      GroupBox 10,0,620,49,"Function List Filter",.FunctionListFilter
      DropListBox 20,21,180,112,Prefixes(),.FunctionNameFilter
      DropListBox 10,56,310,308,FuncNames(),.FunctionName,2 'Lists the (filtered) functions that can be run
      TextBox 10,84,620,308,.TextBox1,2
      CheckBox 360,21,100,14,"Document",.IncludeDoc
      CheckBox 470,21,70,14,"Folder",.IncludeFolder
      CheckBox 550,21,70,14,"Other",.IncludeOther
      TextBox 210,21,130,21,.CustomFilter
      PushButton 330,56,90,21,"Execute",.Execute
      OKButton 530,42,90,21
      PushButton 460,56,170,21,"Continue parent function",.ContinueParentEvent
   End Dialog

   Dim dlg As DevDialog
   dlg.IncludeDoc = (Not pXDoc Is Nothing)
   dlg.IncludeFolder = (Not pXFolder Is Nothing)
   dlg.IncludeOther = True

   DevMenu_ActiveContext(True,pXFolder,pXDoc)

   Dim Result As Integer
   Result=Dialog(dlg)

   Debug.Print("Dialog result: " & Result)

   Select Case Result
      Case -1 ' cancel
         End
      Case 0 ' OK (or X)
         End
      Case 1 ' Execute
         End
      Case 2 ' Continue parent function
   End Select

End Sub

Public Sub DevMenu_ActiveContext(SetActiveVars As Boolean,ByRef pXFolder As CscXFolder, ByRef pXDoc As CscXDocument)
   ' The dialog function processing form events does not allow additional parameters and need a way to get doc/folder
   ' This is essentially the same as using global variables, just a bit more structured and contained
   Static ActiveFolder As CscXFolder
   Static ActiveDoc As CscXDocument

   If SetActiveVars Then
      Set ActiveFolder=pXFolder
      Set ActiveDoc=pXDoc
   Else
      Set pXFolder=ActiveFolder
      Set pXDoc=ActiveDoc
   End If
End Sub

Public Sub DevMenu_DialogInitialize()

End Sub

Public Sub DevMenu_DialogUpdate(AllFunc() As ScriptFunction, ByRef FilteredFunc() As ScriptFunction, ByRef FuncNames() As String, ByRef FilterStr As String, ByRef ScriptStatus As String)
   FilterStr=IIf(DlgValue("FunctionNameFilter")=0,"",IIf(DlgValue("FunctionNameFilter")=1,DlgText("CustomFilter"),DlgText("FunctionNameFilter")))
   FuncNames=DevMenu_FilteredFunctions(AllFunc,FilteredFunc,(DlgValue("IncludeFolder")=1),DlgValue("IncludeDoc"),DlgValue("IncludeOther"),FilterStr)
   ScriptStatus="Total Functions: " & UBound(AllFunc) & ", Filtered Functions: " & UBound(FilteredFunc)
   DlgListBoxArray("FunctionName",FuncNames)
   DlgEnable("CustomFilter", (DlgValue("FunctionNameFilter")=1))
End Sub


Public Function DevMenu_DialogFunc(DlgItem$, Action%, SuppValue?) As Boolean
   Static FilteredFunc() As ScriptFunction
   Static AllFunc() As ScriptFunction
   Static FuncNames() As String
   Static FilterStr As String
   Static CurSF As ScriptFunction
   Static pXDoc As CscXDocument
   Static pXFolder As CscXFolder
   Static ParamStatus As String
   Static ScriptStatus As String

   Select Case Action%
      Case 1 ' Dialog box initialization
         Debug.Print("Dialog: Dialog initialized.  Window handle: " & SuppValue)
         ' Get the folder/doc since they cannot be passed as a parameter to the dialog function
         DevMenu_ActiveContext(False,pXFolder,pXDoc)

         If pXFolder Is Nothing And pXDoc Is Nothing Then ParamStatus="Warning: Neither an XDoc nor XFolder were provided"
         If Not pXFolder Is Nothing And Not pXDoc Is Nothing Then ParamStatus="Both XDoc and XFolder were provided: Functions will use applicable parameter"
         If Not pXFolder Is Nothing And pXDoc Is Nothing Then ParamStatus="FOLDER MODE: Document functions will run on each document in the folder"
         If pXFolder Is Nothing And Not pXDoc Is Nothing Then ParamStatus="DOCUMENT MODE: Folder functions will run on the doc's parent folder"
         If Not pXFolder Is Nothing Then ParamStatus=ParamStatus & vbNewLine & "Folder contains " & pXFolder.DocInfos.Count & " documents."
         If Not pXDoc Is Nothing Then
            ParamStatus=ParamStatus & vbNewLine & "Doc " & pXDoc.IndexInFolder & "/" & pXDoc.ParentFolder.DocInfos.Count & ": " & pXDoc.FileName
         End If

         Dim ParentFunction As String
         On Error Resume Next
         ParentFunction = Mid(CallersLine(1),InStr(1,CallersLine(1),"|")+1,InStr(1,CallersLine(1),"#")-InStr(1,CallersLine(1),"|")-1)
         On Error GoTo 0
         ParamStatus="Launched from parent function: " & ParentFunction & vbNewLine & ParamStatus

         Dim msg As String
         msg="This menu allows for easy execution of design time scripts that work on individual documents, whole folders, or neither.  It will dynamically list all of the project script functions that require no parameters or require only an XDoc/XFolder (with any amount of optional parameters allowed). " & vbNewLine
         msg=msg & vbNewLine & "It should be used from Project Builder in either of two ways:" & vbNewLine
         msg=msg & "1. Providing an XFolder as a parameter, from an event like Batch_Open. Execute the event from the Runtime Script Events button (lightning bolt)." & vbNewLine
         msg=msg & "2. Providing an XDocument as a parameter, from a document level extraction event like Document_BeforeExtract, in a separate class not otherwise used in the project. Execute the event by selecting the class, selecting the document, then extracting the document." & vbNewLine
         msg=msg & vbNewLine & "When an XFolder is provided, Document functions will run on each doc in the folder.  When an XDoc is provided, Folder functions will run on the doc's parent folder.  Not all operations will work when using the parent folder from a document event."

         DlgText("TextBox1",ParamStatus & vbNewLine & vbNewLine & msg)
         'DlgEnable("TextBox1", False)

         AllFunc=ParseScript(Project.ScriptCode,"Project")
         DlgVisible("OK", False) ' OK button must exist to be able to close diaglog via X
         DevMenu_DialogUpdate(AllFunc,FilteredFunc,FuncNames,FilterStr,ScriptStatus)

      Case 2 ' Value changing or button pressed (button press will close dialog unless returning true)
         Debug.Print "Dialog: " & DlgItem & " (" & DlgType(DlgItem) & ") ";
         If InStr(1,DlgType(DlgItem),"Button")=0 Then Debug.Print "value changed to " & SuppValue & " (" & DlgText(DlgItem) & ")." ; : Debug.Print

         Select Case DlgItem
            Case "FunctionNameFilter", "IncludeFolder", "IncludeDoc", "IncludeOther"
               DevMenu_DialogUpdate(AllFunc,FilteredFunc,FuncNames,FilterStr,ScriptStatus)
            Case "FunctionName"
               CurSF=FilteredFunc(DlgValue("FunctionName"))
               DlgText("TextBox1",ParamStatus & vbNewLine & vbNewLine & ScriptStatus & vbNewLine & vbNewLine & CurSF.Content)
            Case "Execute"
               If DlgValue("FunctionName")=-1 Then
                  Return True
               Else
                  CurSF=FilteredFunc(DlgValue("FunctionName"))
                  DevMenu_Execute(CurSF, pXFolder, pXDoc)
               End If
         End Select
      Case 3 ' TextBox or ComboBox text changed
         Debug.Print("Dialog: Text of " & DlgType(DlgItem) & " " & DlgItem & " changed by " & SuppValue & " characters to " & DlgText(DlgItem))
         Select Case DlgItem
            Case "CustomFilter"
               DevMenu_DialogUpdate(AllFunc,FilteredFunc,FuncNames,FilterStr,ScriptStatus)
         End Select
      Case 4 ' Focus changed
         If SuppValue>-1 Then Debug.Print("Dialog: Focus changing from " & DlgName(SuppValue) & " to " & DlgItem)

      Case 5 ' Idle
         Return False ' Prevent further idle actions

      Case 6 ' Function key
         Debug.Print "Dialog: " & IIf(SuppValue And &H100,"Shift-","") & IIf(SuppValue And &H200,"Ctrl-","") & IIf(SuppValue And &H400,"Alt-","") & "F" & (SuppValue And &HFF)
      Case Else
         Debug.Print("Unknown event " & Join(Array(DlgItem,Action,SuppValue),", "))
   End Select
End Function



Public Sub DevMenu_Execute(sf As ScriptFunction, Optional pXFolder As CscXFolder, Optional pXDoc As CscXDocument)
   ' Execute the function using the right parameter based on what is needed and available.
   ' Initially this used Eval to call the function directly, however any unhandled error within eval context causes a crash
   ' and breakpoints are not hit.  Using Eval to get a delegate then invoking it outside of Eval solves these problems.

   ' Use Eval to get a delegate of the function
   Dim EvalStr As String, DelegateVar As Variant
   EvalStr="AddressOf " & sf.Name
   Debug.Print "Delegate Eval: " & EvalStr
   ' Declaring a staticly typed delegate would require a fixed signature, which would not allow for open-ended optional params, or open ended return types
   DelegateVar=Eval(EvalStr)


   ' Invoke delegate with no params
   Dim Result As Variant
   If UBound(sf.Params)=-1 Then
      DynamicInvoke(DelegateVar)
      Exit Sub
   End If

   ' Get the param based on what is needed and what is available
   ' could consider a loop to provide default values or prompt user for additional params
   Dim Param1 As Object
   Select Case Replace(LCase(sf.Params(0).ParamType),"cascadelib.","")
      Case "cscxdocument"
         If Not pXDoc Is Nothing Then
            Set Param1=pXDoc
         Else
            If Not pXFolder Is Nothing AndAlso pXFolder.DocInfos.Count>0 Then
               Debug.Print("Executing document function on each doc in folder.")

               Dim DocIndex As Integer
               For DocIndex=0 To pXFolder.GetTotalDocumentCount()-1
                  Debug.Print("Executing on document " & DocIndex+1 & "/" & pXFolder.DocInfos.Count())
                  DynamicInvoke(DelegateVar,pXFolder.DocInfos(DocIndex).XDocument)
               Next
               Exit Sub

            End If
         End If
      Case "cscxfolder"
         If Not pXFolder Is Nothing Then
            Set Param1=pXFolder
         Else
            If Not pXDoc Is Nothing Then
               Debug.Print("Executing function using parent folder of xdoc (overriding single-document mode and folder access permissions).")
               ' Normally if you go to the parent folder from a doc level event, then back down through the xdocinfos to the xdocs,
               ' that would result in an error saying that it is not currently possible to access documents.
               ' Disabling single doc mode will allow access to the xdocs
               ' These commands are unsupported and have high potential to cause problems: They should never be touched at runtime.
               pXDoc.ParentFolder.SetSingleDocumentMode(False)
               pXDoc.ParentFolder.SetFolderAccessPermission(255)
               Set Param1=pXDoc.ParentFolder
            End If
         End If
   End Select

   If Param1 Is Nothing Then
      Debug.Print("Could not get a valid " & sf.Params(0).ParamType & ", skipping execution.")
   Else
      DynamicInvoke(DelegateVar,Param1)
   End If
End Sub



Public Function DynamicInvoke(DelegateVar As Variant, ParamArray Params() As Variant) As Variant
   ' This allows dynamically invoking a delegate regardless of number of parameters, as well as
   ' helping show where a real error occurs within an invoked function.


   On Error Resume Next
   Dim Result As Variant
   Select Case UBound(Params)
      Case -1
         Result=DelegateVar.Invoke()
      Case 0
         Result=DelegateVar.Invoke(Params(0))
      Case 1
         Result=DelegateVar.Invoke(Params(0), Params(1))
      Case 2
         Result=DelegateVar.Invoke(Params(0), Params(1), Params(2))
      Case Else
         Err.Raise("Define more cases statements to handle more parameters")
   End Select

   If Err.Number=0 Then
      'Debug.Print("Invoked function completed without error.")
   Else
      ' A quirk of debugging invoked functions: IDE will stop execution (+highlight & focus) at the invoke when an error occurs in the invoked function,
      ' however the text of the line causing the actual error will still be changed to red and can be seen from Err.Description.
      ' Instead, we intentionally stop execution here and print the error message to draw attention to this.
      ' If needed, navigate to the stated line number and add a breakpoint.  Then test again to actually stop execution within the invoked function.
      Debug.Print("Real error in invoked function: " & Err.Description)
      Stop ' Refer to Err.Description to see the line number of the real error.
   End If

   ' Output result if it can be converted to string, otherwise resume next
   If CStr(Result)<>"" Then
      Debug.Print("Invoked function result = " & CStr(Result))
   Else
      Debug.Print("Invoked function result = " & TypeName(Result))
   End If

   On Error GoTo 0

   Return Result
End Function


' Functions to test from DevMenu

Public Sub ConvertPdfs(pXFolder As CscXFolder)
   ' Create a sibling folder for tiffs
   Dim fso As New FileSystemObject, ExportPath As String
   ExportPath=fso.GetParentFolderName(fso.GetParentFolderName(pXFolder.FileName)) & "\MultiPageTiff\"
   If Not fso.FolderExists(ExportPath) Then fso.CreateFolder(ExportPath)

   ' Open script window to see progress in intermediate window
   Project.ShowScriptWindow("")

   Dim DocIndex As Integer
   For DocIndex=0 To pXFolder.DocInfos.Count()-1
      ExportDocAsTiff(pXFolder.DocInfos(DocIndex).XDocument, ExportPath, True)
   Next
End Sub

Public Sub ExportDocAsIndividualTiffs(pXDoc As CscXDocument)
   ExportDocAsTiff(pXDoc, GetExportPath(), False)
End Sub

Public Sub ExportDocAsMultipageTiffs(pXDoc As CscXDocument)
   ExportDocAsTiff(pXDoc, GetExportPath(), True)
End Sub

Public Function GetExportPath() As String
   Dim fso As New FileSystemObject
   Dim ExportPath As String
   ExportPath=Project.ScriptVariables("ExportPath")

   If Not fso.FolderExists(ExportPath) Then ExportPath=fso.BuildPath(fso.GetFile(Project.FileName).ParentFolder, "ExportedImages")
   If Not fso.FolderExists(ExportPath) Then fso.CreateFolder(ExportPath)
   Return ExportPath
End Function

Public Sub ExportDocAsTiff(pXDoc As CscXDocument, ExportPath As String, Optional MultiPage As Boolean=True, Optional Bitonal=True)
   Dim fso As New FileSystemObject
   Dim DocName As String, TempPath As String, TiffPath As String
   DocName=fso.GetBaseName(pXDoc.CDoc.SourceFiles(0).FileName)
   If MultiPage Then
      ' Multipage document named by the filename of the first source file
      TiffPath=fso.BuildPath(ExportPath,DocName & ".tif")
      TempPath=TiffPath & ".tmp"
   Else
      ' Create a folder for this document named by the filename of the first source file
      ExportPath=fso.BuildPath(ExportPath,DocName)
      If Not fso.FolderExists(ExportPath) Then fso.CreateFolder(ExportPath)
   End If


   Debug.Print("Exporting " & pXDoc.CDoc.Pages.Count & " pages from document " & pXDoc.IndexInFolder+1 & "/" & pXDoc.ParentFolder.DocInfos.Count & " (" & DocName & ")")
   Dim PageIndex As Integer, img As CscImage, imgformat As CscImageFileFormat
   For PageIndex=0 To pXDoc.CDoc.Pages.Count-1
      If (PageIndex+1) Mod 10 = 0 Then Debug.Print("  Processing page " & PageIndex + 1 & "/" & pXDoc.CDoc.Pages.Count & " from document # " & pXDoc.IndexInFolder+1 & " (" & DocName & ")")

      If MultiPage Then
         If Bitonal Then
            Set img=pXDoc.CDoc.Pages(PageIndex).GetBitonalImage(Project.ColorConversion)
            imgformat=CscImageFileFormat.CscImgFileFormatTIFFFaxG4
         Else
            Set img=pXDoc.CDoc.Pages(PageIndex).GetImage()
            ' It appears that images loaded from PDF might always be treated as color
            imgformat=IIf(img.BitsPerSample=1 And img.SamplesPerPixel=1,CscImageFileFormat.CscImgFileFormatTIFFFaxG4,CscImageFileFormat.CscImgFileFormatTIFFOJPG)
         End If
         img.StgFilterControl(imgformat, CscStgControlOptions.CscStgCtrlTIFFKeepFileOpen, TempPath, 0, 0)
         img.StgFilterControl(imgformat, CscStgControlOptions.CscStgCtrlTIFFKeepExistingPages, TempPath, 0, 0)
         img.Save(TempPath, imgformat)
         img.XResolution
         pXDoc.CDoc.Pages(PageIndex).UnloadImage()
      Else
         pXDoc.CDoc.Pages(PageIndex).GetImage().Save(fso.BuildPath(ExportPath,DocName & "-Page-" & Format(PageIndex+1,"000") & ".tif"),CscImgFileFormatTIFFOJPG)
         pXDoc.CDoc.Pages(PageIndex).UnloadImage()
      End If
   Next

   If MultiPage And pXDoc.CDoc.Pages.Count>0 Then
      ' Close the multipage tiff file that was kept open
      Set img=pXDoc.CDoc.Pages(0).GetImage()
      img.StgFilterControl(CscImageFileFormat.CscImgFileFormatTIFFFaxG4, CscStgControlOptions.CscStgCtrlTIFFCloseFile, TempPath, 0, 0)
      pXDoc.CDoc.Pages(0).UnloadImage()

      ' Delete existing file at destination if needed, then move temp file to destination
      If fso.FileExists(TiffPath) Then fso.DeleteFile(TiffPath)
      fso.MoveFile(TempPath,TiffPath)
   End If

   Debug.Print("Finished " & pXDoc.CDoc.Pages.Count & " pages from document # " & pXDoc.IndexInFolder+1 & " (" & DocName & ")")
End Sub



Private Sub Batch_Open(ByVal pXRootFolder As CASCADELib.CscXFolder)
   ' Invoke DevMenu by testing the Batch_Open function (lightning bolt)
   DevMenu_Dialog(pXRootFolder)

   MsgBox("Post menu")
End Sub


'========  START DEV EXPORT ========
Public Function ClassHierarchy(KtmClass As CscClass) As String
   ' Given TargetClass, returns Baseclass\subclass\(etc...)\TargetClass\

   Dim CurClass As CscClass, Result As String
   Set CurClass = KtmClass

   While Not CurClass.ParentClass Is Nothing
      Result=CurClass.Name & "\" & Result
      Set CurClass = CurClass.ParentClass
   Wend
   Result=CurClass.Name & "\" & Result
   Return Result
End Function

Public Sub CreateClassFolders(ByVal BaseFolder As String, Optional KtmClass As CscClass=Nothing)
   ' Creates folders in BaseFolder matching the project class structure

   Dim SubClasses As CscClasses
   If KtmClass Is Nothing Then
      ' Start with the project class, but don't create a folder
      Set KtmClass = Project.RootClass
      Set SubClasses = Project.BaseClasses
   Else
      ' Create folder for this class and become the new base folder
      Dim fso As New Scripting.FileSystemObject, NewBase As String
      BaseFolder=fso.BuildPath(BaseFolder,KtmClass.Name)
      If Not fso.FolderExists(BaseFolder) Then
         fso.CreateFolder(BaseFolder)
      End If
      Set SubClasses = KtmClass.SubClasses
   End If

   ' Subclasses
   Dim ClassIndex As Long
   For ClassIndex=1 To SubClasses.Count
      CreateClassFolders(BaseFolder, SubClasses.ItemByIndex(ClassIndex))
   Next
End Sub



Public Sub Dev_ExportScriptAndLocators()
   ' Exports design info (script, locators) to to folders matching the project class structure
   ' Default to \ProjectFolderParent\DevExport\(Class Folders)
   ' Set script variable Dev-Export-BaseFolder to path to override
   ' Set script variable Dev-Export-CopyName-(ClassName) to save a separate named copy of a class script

   ' Make sure you've added the Microsoft Scripting Runtime reference
   Dim fso As New Scripting.FileSystemObject
   Dim ExportFolder As String, ScriptFolder As String, LocatorFolder As String

   ' Either use the provided path or default to the parent of the project folder
   If fso.FolderExists(Project.ScriptVariables("Dev-Export-BaseFolder")) Then
      ExportFolder=Project.ScriptVariables("Dev-Export-BaseFolder")
   Else
      ExportFolder=fso.GetFile(Project.FileName).ParentFolder.ParentFolder.Path & "\DevExport"
   End If

   ' Create folder structure for project classes
   If Not fso.FolderExists(ExportFolder) Then fso.CreateFolder(ExportFolder)
   CreateClassFolders(ExportFolder)

   ' Here we use class index -1 to represent the special case of the project class
   Dim ClassIndex As Long
   For ClassIndex=-1 To Project.ClassCount-1
      Dim KtmClass As CscClass, ClassName As String, ScriptCode As String, ClassPath As String

      ' Get the script of this class
      If ClassIndex=-1 Then
         Set KtmClass=Project.RootClass
         ScriptCode=Project.ScriptCode
      Else
         Set KtmClass=Project.ClassByIndex(ClassIndex)
         ScriptCode=KtmClass.ScriptCode
      End If

      ' TODO: Possibly change to match the naming conventions used in the KTM 6.1.1+ feature to save scripts.

      ' Get the name and file path for the class
      ClassPath = fso.BuildPath(ExportFolder, ClassHierarchy(KtmClass))
      ClassName=IIf(ClassIndex=-1,"Project",KtmClass.Name)

      Dim EmptyScript As Boolean
      EmptyScript=(ScriptCode="Option Explicit" & vbNewLine & vbNewLine & "' Class script: " & ClassName)

      ' Export script to file
      If Not EmptyScript Then
         Dim ScriptFile As TextStream
         Set ScriptFile=fso.CreateTextFile(ClassPath & "\ClassScript-" & ClassName & ".vb",True,False)
         ScriptFile.Write(ScriptCode)
         ScriptFile.Close()
      End If

      ' Save a copy if a name is defined
      Dim CopyName As String
      CopyName=Project.ScriptVariables("Dev-Export-CopyName-" & ClassName)

      If Not CopyName="" Then
         Set ScriptFile=fso.CreateTextFile(ClassPath & "\" & CopyName & ".vb",True,False)
         ScriptFile.Write(ScriptCode)
         ScriptFile.Close()
      End If

      ' Export locators (same as from Project Builder menus)
      Dim FileName As String
      Dim LocatorIndex As Integer
      For LocatorIndex=0 To KtmClass.Locators.Count-1
         If Not KtmClass.Locators.ItemByIndex(LocatorIndex).LocatorMethod Is Nothing Then
            FileName="\" & ClassName & "-" & KtmClass.Locators.ItemByIndex(LocatorIndex).Name & ".loc"
            KtmClass.Locators.ItemByIndex(LocatorIndex).ExportLocatorMethod(ClassPath & FileName, ClassPath)
         End If
      Next
   Next
End Sub
'========  END   DEV EXPORT ========

Public Sub Example_MsgBox()
   MsgBox("Test")
End Sub

Private Sub Document_AfterClassifyXDoc(ByVal pXDoc As CASCADELib.CscXDocument)
   Debug.Print(Logging_StackLine(CallersLine(-1)))
End Sub

Private Sub Document_AfterExtract(ByVal pXDoc As CASCADELib.CscXDocument)
   Debug.Print(Logging_StackLine(CallersLine(-1)))
End Sub

Private Sub Document_AfterProcess(ByVal pXDoc As CASCADELib.CscXDocument)
   Debug.Print(Logging_StackLine(CallersLine(-1)))
End Sub

Private Sub Document_AfterSeparatePages(ByVal pXDoc As CASCADELib.CscXDocument)
   Debug.Print(Logging_StackLine(CallersLine(-1)))
End Sub

Private Sub Document_BeforeClassifyXDoc(ByVal pXDoc As CASCADELib.CscXDocument, ByRef bSkip As Boolean)
   Debug.Print(Logging_StackLine(CallersLine(-1)))
End Sub

Private Sub Document_BeforeExtract(ByVal pXDoc As CASCADELib.CscXDocument)
   Debug.Print(Logging_StackLine(CallersLine(-1)))
End Sub

Private Sub Document_BeforeSeparatePages(ByVal pXDoc As CASCADELib.CscXDocument, ByRef bSkip As Boolean)
   Debug.Print(Logging_StackLine(CallersLine(-1)))
End Sub

Private Sub Document_BeforeTDS(ByVal pXDoc As CASCADELib.CscXDocument, ByRef bSkip As Boolean)
   Debug.Print(Logging_StackLine(CallersLine(-1)))
End Sub
