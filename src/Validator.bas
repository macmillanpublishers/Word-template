Attribute VB_Name = "Validator"
' =============================================================================
'     BOOKMAKER VALIDATOR
' =============================================================================

' By Erica Warren - erica.warren@macmillan.com
'
' ===== USE ===================================================================
' To validate/fix Word manuscripts automatically before createing egalleys.
'
' ===== DEPENDENCIES ==========================================================
' + !! PC ONLY !!!  Not testing anything on Mac.
' + Powershell script that calls this macro is looking for `Validator.Launch`
'   with parameters (1) FilePath and (2) LogPath.
' + ALL `MsgBox` called in ANY procedure MUST return correct default for this
'   macro to continue.


' ===== Global Declarations ===================================================
Option Explicit

' For error checking:
Private Const strValidator As String = "Bookmaker.Validator."
Private lngCleanupCount As Long
' Create path for alert file in same dir as ACTIVE doc (NOT ThisDocument)
Private strAlertPath As String
' Store style check pass/fail values in this json
Private strJsonPath As String
' Ditto but for log file
Private strLogPath As String
' Name of originally called procedure, to write to logs
Private strProcName As String

Private StartTime As Double
Private SecondsElapsed As Double

' store info from `book_info.json` file
Private dictBookInfo As Dictionary
' store initial paragraph loop info
Private dictStyles As Dictionary
' store acceptable heading styles
Private dictHeadings As Dictionary
' store macmillan styles based on other mac styles
Private c_dictBaseStyle As Dictionary
' store style-to-section conversion
Private dictSections As Dictionary
' also path to write alerts to
Private strAlertFile As String

Private strVersion As String

' A bunch of styles we'll need.
Private Const strPageBreak As String = "Page Break (pb)"
Private Const strSectionBreak As String = "Section Break (sbr)"
Private Const strChapTitle As String = "Chap Title (ct)"
Private Const strChapNumber As String = "Chap Number (cn)"
Private Const strChapNonprinting As String = "Chap Title Nonprinting (ctnp)"
Private Const strIsbnStyle As String = "span isbn (ISBN)"
Private Const strBookTitle As String = "Titlepage Book Title (tit)"
Private Const strAuthorName As String = "Titlepage Author Name (au)"
Private Const strCopyright As String = "Copyright Text single space (crtx)"
Private Const strCopyright2 As String = "Copyright Text double space (crtxd)"
Private Const strBodyStyle As String = "Text - Standard (tx)"
Private Const strFmEpiText As String = "FM Epigraph - non-verse (fmepi)"
Private Const strFmEpiVerse As String = "FM Epigraph - verse (fmepiv)"
Private Const strFmEpiSource As String = "FM Epigraph Source (fmeps)"
Private Const strHalftitle As String = "Halftitle Book Title (htit)"
Private Const strPartTitle As String = "Part Title (pt)"
Private Const strPartNumber As String = "Part Number (pn)"
Private Const strFmHead As String = "FM Head (fmh)"
Private Const strFmHeadAlt As String = "Fm Head ALT (afmh)"
Private Const strFmTitle As String = "FM Title (fmt)"
Private Const strBmHead As String = "BM Head (bmh)"
Private Const strBmHeadAlt As String = "BM Head ALT (abmh)"
Private Const strBmTitle As String = "BM Title (bmt)"
Private Const strIllustrationHolder As String = "Illustration holder (ill)"
Private Const c_strFsqHead As String = "Front Sales Quote Head (fsqh)"
Private Const c_strAppHead As String = "Appendix Head (aph)"
Private Const c_strAtaHead As String = "About Author Text Head (atah)"
Private Const c_strSeriesHead As String = "Series Page Heading (sh)"
Private Const c_strAdCardHead As String = "Ad Card Main Head (acmh)"
Private Const c_strRecipeHead As String = "Recipe Head (rh)"
Private Const c_strSubRecipeHead As String = "Sub-Recipe Head (srh)"
Private Const c_strRecipeVarHead As String = "Recipe Var Head (rvh)"
Private Const c_strPoemTitle As String = "Poem Title (vt)"
Private Const c_strFmHeadNonprinting As String = "FM Head Nonprinting (fmhnp)"
Private Const c_strBmHeadNonprinting As String = "BM Head Nonprinting (bmhnp)"

Private Enum BookInfo
  bk_Title = 1
  bk_Authors = 2
  bk_ISBN = 3
End Enum

Private Enum SectionsJson
  j_text = 1
  j_style = 2
End Enum

' ===== Enumerations ==========================================================
Public Enum ValidatorError
  err_ValErrGeneral = 30000
  err_TestsFailed = 30001
  err_PathInvalid = 30003
  err_NoPassKey = 30004
  err_StartupFail = 30005
  err_ValidatorMainFail = 30006
  err_IsbnMainFail = 30007
End Enum

' ===== IsbnSearch ============================================================
' Searches doc for ISBN, called from Powershell.

' PUBLIC because needs to be called independently from powershell if file name
' doesn't include ISBN. Optional FilePath is for passing doc path from PS.

Public Function IsbnSearch(FilePath As String, Optional LogFile As String) _
  As String
  On Error GoTo IsbnSearchError
  strProcName = strValidator & "IsbnSearch"
  
' Startup checks.
  If ValidatorStartup(FilePath, LogFile) = False Then
    Err.Raise ValidatorError.err_StartupFail
  End If

' Main procedure
  Dim strIsbn As String
  strIsbn = IsbnMain(FilePath)

'' Returned false doesn't necessarily mean a problem happened. Maybe just didn't
'' find anything:
'  If strIsbn = vbNullString Then
'    Err.Raise ValidatorError.err_IsbnMainFail
'  End If
  
  IsbnSearch = strIsbn

' Various cleanup stuff, but don't `End` because macro needs to return a value.
  Call ValidatorExit(RunCleanup:=True, EndMacro:=False)
  
  Exit Function

IsbnSearchError:
  Err.Source = strValidator & "IsbnSearch"
  Select Case Err.Number
    Case ValidatorError.err_StartupFail
      Err.Description = strValidator & "ValidatorStartup failed to complete."
    Case ValidatorError.err_IsbnMainFail
      Err.Description = strValidator & "IsbnMain failed to complete."
  End Select
  Call ValidatorExit(RunCleanup:=False)
End Function


' ===== Launch ================================================================
' Set up error checking, suppress alerts.
' Note: (1) ps1 script that calls this will (a) verify file is a Word
' doc, and (b) fix file name so it doesn't include spaces or special characters
' (2) Must handle ALL errors (write to ALERT file in same dir), (3) Must set
' `DisplayAlerts = False`, and be sure ALL msgbox return correct default.

Public Sub Launch(FilePath As String, Optional LogPath As String)
  On Error GoTo LaunchError
  strProcName = strValidator & "Launch"
' Startup checks
  If ValidatorStartup(FilePath, LogPath) = False Then
    Err.Raise ValidatorError.err_StartupFail
  End If

' Main procedure
  If ValidatorMain(FilePath) = True Then
    SecondsElapsed = Round(Timer - StartTime, 2)
    DebugPrint "`Main` function complete: " & SecondsElapsed
  Else
    Err.Raise ValidatorError.err_ValidatorMainFail
  End If

' Various cleanup stuff, including `End` all code execution.
  Call ValidatorExit(RunCleanup:=True, EndMacro:=False)
  
  Exit Sub

LaunchError:
  Err.Source = strValidator & "Launch"
  Select Case Err.Number
    Case ValidatorError.err_StartupFail
      Err.Description = strValidator & "ValidatorStartup failed to complete."
    Case ValidatorError.err_ValidatorMainFail
      Err.Description = strValidator & "ValidatorMain failed to complete."
  End Select
  Call ValidatorExit(RunCleanup:=False)
End Sub

' ===== ValidatorStartup ======================================================
' File and log paths by ref because we might change them.

Private Function ValidatorStartup(ByRef StartupFilePath As String, ByRef _
  StartupLogPath As String) As Boolean
  On Error GoTo ValidatorStartupError
  ValidatorStartup = False
  StartTime = Timer
  Application.DisplayAlerts = wdAlertsNone
'  Application.ScreenUpdating = False

' Fix path separators in file path names
  Dim strFilePath As String
  Dim strLogPath As String
  strFilePath = FilePathCleanup(StartupFilePath)
  strLogPath = FilePathCleanup(StartupLogPath)
  
' set global variable for path to write alert messages to, returns False if
' FilePath doesn't exist or point to a real file.
  If SetOutputPaths(strFilePath, strLogPath) = False Then
    Err.Raise err_PathInvalid
  End If
  SecondsElapsed = Round(Timer - StartTime, 2)
  DebugPrint "Output paths set: " & SecondsElapsed

  ValidatorStartup = True
  Exit Function
  
ValidatorStartupError:
  Err.Source = strValidator & "ValidatorStartup"
' Have to assume here error may occur before we can access general error
' checker, so do everything in this module.
  Select Case Err.Number
    Case err_PathInvalid
      Err.Description = "The string passed for the `FilePath` argument, " & _
        Chr(34) & StartupFilePath & Chr(34) & ", does not point to a valid file."
  End Select
  Call ValidatorExit(RunCleanup:=False)
End Function

' ===== ValidatorExit ======================================================
' Always run last. In fact, it ends ALL macro execution so by definition it'll
' be last! If Err object is not 0, will write an ALERT. `EndMacro` should be
' set to False if you want code execution to return to the calling procedure.

Public Sub ValidatorExit(Optional RunCleanup As Boolean = True, Optional _
  EndMacro As Boolean = True)
' Global variable counter in case error throw before we reset On Error.
' More than 1 is an error, but letting it run a few times to capture more data
  lngCleanupCount = lngCleanupCount + 1
  DebugPrint "ValidatorExit: " & lngCleanupCount
  If lngCleanupCount > 3 Then GoTo ValidatorExitError

' NOTE!! Must get Err object values before setting new On Error statement.
' Did macro complete correctly?
  Dim blnCompleted As Boolean
  If Err.Number = 0 Then
    blnCompleted = True
  Else
    blnCompleted = False
    Call WriteAlert(False)
  End If

' Now we can reset On Error (which includes Err.Clear)
  On Error GoTo ValidatorExitError
  
' Do we know if doc is styled?
  Dim blnStyled As Boolean
  If RunCleanup = True Then
    blnStyled = ValidatorCleanup(blnCompleted)
    Call JsonToLog
  Else
    blnStyled = False
  End If

' Only save doc if macro completed w/o error AND correctly styled
' And not if we're debugging
  Dim saveValue As Boolean
  If Environ("VbaDebug") <> "True" Then
    If blnCompleted = True And blnStyled = True Then
      saveValue = True
    Else
      saveValue = False
    End If
  End If
  
' Close all open documents
  Dim objDoc As Document
  Dim strExt As String
  For Each objDoc In Documents
  ' don't close any templates, might be running code.
    strExt = VBA.Right(objDoc.Name, InStr(StrReverse(objDoc.Name), "."))
    If strExt <> ".dotm" And saveValue = True Then
      objDoc.Save
    End If
  Next objDoc
  
' DON'T `Exit Sub` before this - we want it to `End` no matter what.
ValidatorExitError:
  If Err.Number <> 0 Then
    Err.Source = strValidator & "ValidatorExit"
    Call WriteAlert(False)
  End If
  
  Application.DisplayAlerts = wdAlertsAll
  Application.ScreenUpdating = True
  Application.ScreenRefresh

' Timer End
  SecondsElapsed = Round(Timer - StartTime, 2)
  DebugPrint "This code ran successfully in " & SecondsElapsed & " seconds"

  If EndMacro = True Then
    On Error GoTo 0
  ' Stops ALL code execution (might be called to cleanup after error).
    End
  End If
End Sub

' ===== FilePathCleanup =======================================================
' Ensures path separators are correct for a file path or directory. At some
' point figure out how to include OSX paths?

Private Function FilePathCleanup(ByRef FullFilePath As String) As String
  On Error GoTo FilePathCleanupError
  Dim strReturn As String
  strReturn = FullFilePath
  If InStr(strReturn, "/") > 0 Then
    strReturn = VBA.Replace(strReturn, "/", Application.PathSeparator)
  End If
'  DebugPrint strReturn
  FilePathCleanup = strReturn
  Exit Function
  
FilePathCleanupError:
  Err.Source = strValidator & "FilePathCleanup"
  Call ValidatorExit(RunCleanup:=False)
End Function


' ===== SetOutputPaths ========================================================
' Set local path to write Alerts (i.e., unhandled errors). Must declare private
' global variable up top! On server, tries to write to same path as the file
' passed to Launch, if fails defaults to `validator_tmp`.

Private Function SetOutputPaths(origPath As String, origLogPath As String) As Boolean
  On Error GoTo SetOutputPathsError
  Dim strDir As String
  Dim strFile As String
  Dim lngSep As Long
  

  ' Validate file path. `Dir("")` returns first file or dir in default Templates
  ' dir so we have to check for null string AND if file exists...
  If origPath <> vbNullString And Dir(origPath) <> vbNullString Then
    ' File exists (thus, directory exists too)
    SetOutputPaths = True
    ' Separate directory from file name
    lngSep = InStrRev(origPath, Application.PathSeparator)
    strDir = VBA.Left(origPath, lngSep)  ' includes trailing separator
    strFile = VBA.Right(origPath, Len(origPath) - lngSep)
'    DebugPrint strDir & " | " & strFile
  
  ' If file DOESN'T exist, set defaults
  Else
    SetOutputPaths = False
    Dim strLocalUser As String
    ' If we're on server, use validator default location
    strLocalUser = Environ("USERNAME")
    If strLocalUser = "padwoadmin" Then ' we're on the server
      strDir = "S:/validator_tmp/"
    ' If not, just use desktop
    Else
      strDir = Environ("USERPROFILE") & Application.PathSeparator & "Desktop" _
        & Application.PathSeparator
    End If
  End If
  
  ' build full alert file name
  strFile = "ALERT_" & strFile & "_" & Format(Date, "yyyy-mm-dd") & ".txt"
'  DebugPrint strFile
  
  ' combine path & file name!
  ' this is a global var that WriteAlert function can access directly.
  strAlertPath = strDir & strFile
'  DebugPrint strAlertPath

' Create path to JSON to store test results. Need different names, since the
' powershell reads them too.
  Dim strMacroType As String
  Select Case strProcName
    Case strValidator & "Launch"
      strMacroType = "style"
    Case strValidator & "IsbnSearch"
      strMacroType = "isbn"
  End Select
  strJsonPath = strDir & strMacroType & "_check.json"
  
  ' delete if it exists already (don't want old test results)
  If strJsonPath <> vbNullString And Dir(strJsonPath) <> vbNullString Then
    Kill strJsonPath
  End If

  ' Also verify log file. Could add more error handling later but for now
  ' just trusting that will be created by calling .ps1 script
  strLogPath = origLogPath
  Exit Function
  
SetOutputPathsError:
  Err.Source = strValidator & "SetOutputPaths"
  Call ValidatorExit(RunCleanup:=False)
End Function


' ===== WriteAlert ============================================================
' First intended as last resort if refs are missing, but maybe Err is always
' returned by ErrorChecker (or it's passed ByRef, so it's just updated), and
' we always write the Alert from the primary project. Then different projects
' can handle where to write alerts differently.

' Note `strAlertPath` is a private global variable that needs to be created
' before this is run.

Private Sub WriteAlert(Optional blnEnd As Boolean = True)
  ' Create log message
  Dim strAlert As String
  strAlert = "=========================================" & vbNewLine & _
    Now & " | " & Err.Source & vbNewLine & _
    Err.Number & ": " & Err.Description & vbNewLine

  ' Append message to log file
  Dim FileNum As Long
  FileNum = FreeFile()
  Open strAlertPath For Append As #FileNum
  Print #FileNum, strAlert
  Close #FileNum

'  SecondsElapsed = Round(Timer - StartTime, 2)
'  DebugPrint "WriteAlert: " & strAlert & SecondsElapsed
  
  ' Optional: stops ALL code.
  If blnEnd = True Then
    If Not activeDoc Is Nothing Then
      Set activeDoc = Nothing
    End If
    Application.DisplayAlerts = wdAlertsAll
    Application.ScreenUpdating = True
    Application.ScreenRefresh
    On Error GoTo 0
    End
  End If
End Sub


' ===== IsbnMain ==============================================================
' Implementation of IsbnCheck that returns a string.

Private Function IsbnMain(FilePath As String) As String
  On Error GoTo IsbnMainError
  
' ----- START JSON FILE -------------------------------------------------------
  Call ClassHelpers.AddToJson(strJsonPath, "completed", False)
  
' Make sure relevant file exists, is open
  If Utils.IsOpen(FilePath) = False Then
    Documents.Open FilePath
  End If

' Set reference to correct document
  Set activeDoc = Documents(FilePath)

' ----- Start checking document -----------------------------------------------
  Dim strKey As String
  Dim dictTests As Dictionary
  
' ----- Various startup checks, including doc password check ------------------
  strKey = "initialize"
  Set dictTests = Reports.ReportsStartup(DocPath:=FilePath, _
    AlertPath:=strAlertPath, BookInfoReqd:=False)
  ' Can't use QuitIfFailed:=False below, IsbnCheck (below) doesn't work
  ' if a doc is password protected. Could probably fix but might be a
  ' lot of work.
  Call ReturnDict(strKey, dictTests)

' ----- Search for ISBNs ------------------------------------------------------
  Dim dictIsbn As Dictionary
  strKey = "isbn"
  Set dictIsbn = IsbnCheck(AddFromJson:=False)
  Call ReturnDict(strKey, dictIsbn)

  If Not dictIsbn Is Nothing Then
'  ' JsonToLog function expects this format
'    Dim dictOutput As Dictionary
'    Set dictOutput = ClassHelpers.NewDictionary
'    dictOutput.Add "isbn", dictIsbn
'    Call ClassHelpers.WriteJson(strJsonPath, dictOutput)

  ' If ISBNs were found, they will be in the "list" element
    If dictIsbn.Exists("list") = True Then
    ' Reduce array elements to a comma-delimited string
      IsbnMain = Join(SourceArray:=dictIsbn.Item("list"), Delimiter:=",")
    Else
      IsbnMain = vbNullString
    End If
  Else
    Err.Raise ValidatorError.err_IsbnMainFail
  End If



  Exit Function
  
IsbnMainError:
  Err.Source = strValidator & "IsbnMain"
  If ErrorChecker(Err, FilePath) = False Then
    Resume
  Else
    Call ValidatorExit
  End If
End Function


' ===== Main ==================================================================
' Once we know we've got the correct references set up, we can build our macro.
' DocPath exists and is a Word file already validated.

Private Function ValidatorMain(DocPath As String) As Boolean
  On Error GoTo ValidatorMainError
  ' Set up variables to store test results
  Dim strKey As String
  Dim blnPass As Boolean
  Dim dictTests As Dictionary

' ----- START JSON FILE -------------------------------------------------------
  Call ClassHelpers.AddToJson(strJsonPath, "completed", False)
  
' NOTE! Each procedure called returns a dictionary with results of various
' tests. Each will include a "pass" key--a value of "False" means that the
' validator should NOT continue (checked in `ReturnDict` sub).
  
' ----- INITIALIZE ------------------------------------------------------------
  strKey = "initialize"
  Set dictTests = Reports.ReportsStartup(DocPath, strAlertPath)
  Call ReturnDict(strKey, dictTests)

' *****************************************************************************
'       ALWAYS CHECK ISBN and STYLES
' *****************************************************************************

' ----- OVERALL STYLE CHECKS --------------------------------------------------
  strKey = "styled"
  Set dictTests = Reports.StyleCheck()
  Call ReturnDict(strKey, dictTests)

' ----- ISBN VALIDATION -------------------------------------------------------
' Delete any ISBNS and replace with ISBN from book_info.json
  strKey = "isbn"
  Set dictTests = Reports.IsbnCheck
  Call ReturnDict(strKey, dictTests)

' *****************************************************************************
'       CONTINUE IF MS IS STYLED
' *****************************************************************************

' ----- TITLEPAGE VALIDATION --------------------------------------------------
  strKey = "titlepage"
  Set dictTests = Reports.TitlepageCheck
  Call ReturnDict(strKey, dictTests)

' ----- SECTION TAGGING -------------------------------------------------------
  strKey = "sections"
  Set dictTests = Reports.SectionCheck
  Call ReturnDict(strKey, dictTests)
  
' ----- HEADING VALIDATION ----------------------------------------------------
  strKey = "headings"
  Set dictTests = Reports.HeadingCheck
  Call ReturnDict(strKey, dictTests)
  
'' ----- ILLUSTRATION VALIDATION -----------------------------------------------
  strKey = "illustrations"
  Set dictTests = Reports.IllustrationCheck
  Call ReturnDict(strKey, dictTests)
  
'' ----- ENDNOTE UNLINKING ----------------------------------------------------
  strKey = "endnotes"
  Set dictTests = Endnotes.EndnoteCheck
  Call ReturnDict(strKey, dictTests)

' ----- RUN CLEANUP MACRO -----------------------------------------------------
' To do: convert to function that returns dictionary of test results
  strKey = "cleanupMacro"
  Set dictTests = CleanupMacro.MacmillanManuscriptCleanup
  Call ReturnDict(strKey, dictTests)

' ----- RUN CHAR STYLES MACRO -------------------------------------------------
' To do: convert to function that returns dictionary of test results
  strKey = "characterStyles"
  Set dictTests = CharacterStyles.MacmillanCharStyles
  Call ReturnDict(strKey, dictTests)
''
  Set dictTests = Nothing
  
  ValidatorMain = True
  Exit Function
ValidatorMainError:
  Err.Source = strValidator & "ValidatorMain"
  Select Case Err.Number
    Case ValidatorError.err_TestsFailed
      Err.Description = "The test dictionary for `" & strKey & "` returned empty."
      Call ValidatorExit
    Case ValidatorError.err_NoPassKey
      Err.Description = strKey & " dictionary has no `pass` key."
      Call ValidatorExit
    Case Else
      If MacroHelpers.ErrorChecker(Err) = False Then
        Resume
      Else
        Call ValidatorExit
      End If
  End Select
  
End Function


' ===== ValidatorCleanup ======================================================
' Cleanup items. Returns if doc is styled or not.

Private Function ValidatorCleanup(CompleteSuccess As Boolean) As Boolean
  On Error GoTo ValidatorCleanupError

' BUG: Can't run zz_clearFind if doc is pw protected
' run before set activeDoc = Nothing
  MacroHelpers.zz_clearFind

' just to keep things tidy
  If Not activeDoc Is Nothing Then
    Set activeDoc = Nothing
  End If
  
' Read style_check.json into dictionary in order to access styled info
  Dim dictJson As Dictionary
  Set dictJson = ClassHelpers.ReadJson(strJsonPath)

' Check if all subsections passed
  Dim key1 As Variant
  For Each key1 In dictJson.Keys
    If VBA.IsObject(dictJson(key1)) = True Then
    ' styled pass = False if not styled, which is as designed
      If key1 <> "styled" And dictJson(key1).Item("pass") = False Then
        CompleteSuccess = False
        Exit For
      End If
    End If
  Next key1

' Write our final element to JSON file
  Call AddToJson(strJsonPath, "completed", CompleteSuccess)
  
' Determine if doc is styled or not. If errored before checking styles, item
' won't exist.
  ValidatorCleanup = False
  If Not dictJson Is Nothing Then
    If dictJson.Exists("styled") = True Then
      If dictJson("styled").Exists("pass") = True Then
        ValidatorCleanup = dictJson("styled")("pass")
      End If
    End If
  End If
  Exit Function

ValidatorCleanupError:
  Err.Source = strValidator & "ValidatorCleanup"
' This gets called from ValidatorExit, so skip here and just write alert.
  Call WriteAlert

End Function


' ===== JsonToLog =============================================================
' Converts `style_check.json` to human-readable log entry, and writes to log.

Public Sub JsonToLog()
  On Error GoTo JsonToLogError
' Determine how many seconds code took to run
  SecondsElapsed = Round(Timer - StartTime, 2)
  Call AddToJson(strJsonPath, "elaspedTime", SecondsElapsed)

' Write log entry from JSON values
' Should always be there (see previous line)
  If Utils.IsItThere(strJsonPath) = True Then

  Dim jsonDict As Dictionary
    Set jsonDict = ClassHelpers.ReadJson(strJsonPath)
    
    Dim strLog As String  ' string to write to log
  ' Following Matt's formatting for other scripts
    strLog = Format(Now, "yyyy-mm-dd hh:mm:ss AMPM") & "   : " & strProcName _
       & " -- results:" & vbNewLine
    
  ' Loop through `style_check.json` and write to log. Can add more detailed info
  ' in the future.
  
  ' Also don't stress now, but in future could write more generic dict-to-json
  ' function by breaking into multiple functions that call each other:
  ' Value is an array (return all items in comma-delineated string - REDUCE!)
  ' Value is an object (call this function again!)
  ' Value is neither (thus, number, string, boolean) - just write to string.
    Dim strKey1 As Variant
    Dim strKey2 As Variant
    Dim colValues As Collection
    Dim A As Long
    
  ' Anyway, loop through json data and build string to write to log
    With jsonDict
      For Each strKey1 In .Keys
        strLog = strLog & strKey1 & ": "
      ' Value may be another dictionary/object
        If VBA.IsObject(.Item(strKey1)) = True Then
          strLog = strLog & vbNewLine
        ' loop through THIS dictionary
          For Each strKey2 In .Item(strKey1).Keys
            strLog = strLog & vbTab & strKey2 & ": "
          ' Value at this level may be a Collection (how Dictionary class returns
          ' an array from JSON)
            If strKey2 = "list" Then
              If VBA.IsObject(.Item(strKey1).Item(strKey2)) = True Then
                strLog = strLog & vbNewLine & vbTab & vbTab
                Set colValues = .Item(strKey1).Item(strKey2)
              ' Loop through collection, write values
                For A = 1 To colValues.Count
                  If A <> 1 Then
                    ' add comma and space between values
                    strLog = strLog & ", "
                  End If
                  strLog = strLog & colValues(A)
                Next A
              End If
            Else
            ' Pretty sure it's something we can convert to string directly
              strLog = strLog & .Item(strKey1).Item(strKey2)
            End If
            strLog = strLog & vbNewLine
          Next strKey2
        Else
          strLog = strLog & .Item(strKey1) & vbNewLine
        End If
  '      strLog = strLog & vbNewLine
      Next strKey1
    End With
    strLog = strLog & vbNewLine
  '  DebugPrint strLog
    
  ' Write string to log file, which should have been set earlier!
  '  DebugPrint strLogPath
    Call Utils.AppendTextFile(strLogPath, strLog)
  End If
  
  Exit Sub

JsonToLogError:
' Call WriteAlert not ValidatorCleanup, because called from ValidatorCleanup
  Err.Source = strValidator & "JsonToLog"
  Call WriteAlert(blnEnd:=True)
  
End Sub


' ===== ReturnDict ============================================================
' Process dictionary returned from reports section. If the procedure that made
' the dictionary returns "pass":False then whole macro will quit unless you
' set QuitIfFailed:=False when calling this sub.

Private Sub ReturnDict(SectionKey As String, TestDict As Dictionary, _
  Optional QuitIfFailed As Boolean = True)
  On Error GoTo ReturnDictError
  If TestDict Is Nothing Then
    Err.Raise ValidatorError.err_TestsFailed
  Else
    If TestDict.Exists("pass") = True Then
      ' write tests to JSON file
      Call ClassHelpers.AddToJson(strJsonPath, SectionKey, TestDict)
      If TestDict("pass") = False Then
        If QuitIfFailed = True Then
          SecondsElapsed = Round(Timer - StartTime, 2)
          DebugPrint SectionKey & " complete: " & SecondsElapsed
        ' ValidatorCleanup will end code execution
          Call ValidatorExit
        End If
      End If
    Else
      Err.Raise ValidatorError.err_NoPassKey
    End If
  End If

  SecondsElapsed = Round(Timer - StartTime, 2)
  DebugPrint SectionKey & " complete: " & SecondsElapsed
  
  Exit Sub
  
ReturnDictError:
  Err.Source = strValidator & "ReturnDict"
  Select Case Err.Number
    Case ValidatorError.err_TestsFailed
      Err.Description = "The test dictionary for `" & SectionKey & "` returned empty."
      Call ValidatorExit
    Case ValidatorError.err_NoPassKey
      Err.Description = SectionKey & " dictionary has no `pass` key."
      Call ValidatorExit
    Case Else
      If MacroHelpers.ErrorChecker(Err) = False Then
        Resume
      Else
        Call ValidatorExit
      End If
  End Select

End Sub

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'          FOR TESTING
' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Public Sub ValidatorTest()
'' to simulate being called by ps1
  On Error GoTo TestError
  Dim strFile As String
  Dim strDir As String
  
  strDir = Environ("USERPROFILE") & "\Desktop\validator\"
  strFile = "_validatortest_orig"
'  strFile = "EdwardVIIManuscriptReadyforComp"


  Call Validator.Launch(strDir & strFile & ".docx", _
  strDir & "LOG_" & strFile & ".txt")
  Exit Sub

TestError:
  DebugPrint Err.Number & ": " & Err.Description
End Sub


Public Sub IsbnTest()
'' to simulate being called by ps1
  On Error GoTo TestError
  Dim strDir As String
  Dim strLog As String
  Dim strThisFile As String
  Dim strReturnedIsbn As String
  Dim strFile As String
  Dim A As Long

  
  strDir = Environ("USERPROFILE") & "\Desktop\validator\"

'  strFile = "09Chaput_UNSTYLED_inText_styles-added"
'  strFile = "9781627790031_The_Book_of_Shadows_FINAL"
strFile = "validatortest_orig"
  strLog = strDir & strFile & ".txt"
  strThisFile = strDir & strFile & ".docx"
  strReturnedIsbn = Validator.IsbnSearch(strThisFile, strLog)

  If strReturnedIsbn = vbNullString Then
    DebugPrint "No Isbn found"
  Else
    DebugPrint "Found Isbns: " & strReturnedIsbn
  End If

  Exit Sub

TestError:
  DebugPrint Err.Number & ": " & Err.Description
End Sub

' ===== ReportsStartup ========================================================
' Set some global vars, check some things. Probably should be an Initialize
' event at some point?

Private Function ReportsStartup(DocPath As String, AlertPath As String, Optional _
  BookInfoReqd As Boolean = True) As Dictionary
  On Error GoTo ReportsStartupError
  
' Store test data
  Dim dictReturn As Dictionary
  Set dictReturn = New Dictionary
  dictReturn.Add "pass", False

' Get this first, in case we have an error early:
  strAlertFile = AlertPath

' The .ps1 that calls this macro also opens the file, so should already be
' part of the Documents collection, but we'll check anyway.
  If Utils.IsOpen(DocPath) = False Then
    Documents.Open (DocPath)
  End If

' Set object variable (global scope, so other procedures can use)
  Set activeDoc = Documents(DocPath)
  
' Check for `book_info.json` file, read into global dictionary variable
  If BookInfoReqd = True Then
    Dim strInfoPath As String
    strInfoPath = activeDoc.Path & Application.PathSeparator & "book_info.json"
    If Utils.IsItThere(strInfoPath) = True Then
      Set dictBookInfo = ClassHelpers.ReadJson(strInfoPath)
    Else
      Err.Raise MacError.err_FileNotThere
    End If
  Else
    Set dictBookInfo = New Dictionary
  End If

' Check that doc is not password protected, if so exit
' value will be written to JSON, validator will have to report to user.
  If activeDoc.ProtectionType <> wdNoProtection Then
    dictReturn.Add "password_protected", True
    Set ReportsStartup = dictReturn
    Exit Function
  Else
    dictReturn.Add "password_protected", False
  End If

' Turn off Track Changes
  activeDoc.TrackRevisions = False
  
' Check for placeholders (from failed macros), remove if found
  Dim blnPlaceholders As Boolean
  With activeDoc.Range.Find
    MacroHelpers.zz_clearFind
    .MatchWildcards = True
    .Text = "([`|]{1,2}[0-9A-Z][`|]{1,2}){1,}"
    .Replacement.Text = ""
    .Execute Replace:=wdReplaceAll
    
    blnPlaceholders = .Found
  End With
  dictReturn.Add "cleanup_placeholders_found", blnPlaceholders
  
  If Not dictBookInfo Is Nothing Then
    dictReturn.Item("pass") = True
  End If
  
  OldStartStyles
  
  Set ReportsStartup = dictReturn
  
  Exit Function
ReportsStartupError:
  Err.Source = strReports & "ReportsStartup"
  If ErrorChecker(Err, strInfoPath) = False Then
    Resume
  Else
    Call Reports.ReportsTerminate
  End If
End Function


' ===== StyleVersion ==========================================================
' Checks if we have a style version and if so, adds it as a processing instruction

Private Function OldStartStyles() As Boolean
' private module-level variable
  strVersion = Reports.GetStyleVersion()
  If strVersion = vbNullString Then
    OldStartStyles = True
  Else
    OldStartStyles = False
    ' add as Processing instruction
    Dim rngEnd As Range
    Set rngEnd = activeDoc.Range
    rngEnd.InsertAfter vbNewLine & "VERSION: " & strVersion
    rngEnd.Collapse wdCollapseEnd
    rngEnd.Style = "Bookmaker Processing Instruction (bpi)"
  End If

End Function

' ===== ReportsTerminate ======================================================
' Things to do if we have to terminate the macro early due to an error. To be
' called if ErrorChecker returns false. Again, some day would work better as
' a class with a legit Class_Terminate procedure.

Private Sub ReportsTerminate()
  Dim lngErrNumber As Long
  Dim strErrDescription As String
  Dim strErrSource As String
  
  lngErrNumber = Err.Number
  strErrDescription = Err.Description
  strErrSource = Err.Source

' Get current Err values before new `On Error` (which clear Err object)
  On Error GoTo ReportsTerminateError

' Write error file to active doc dir
  If strAlertFile = vbNullString Then
    If Not activeDoc Is Nothing Then
      strAlertFile = activeDoc.Path
    Else
      strAlertFile = activeDoc.Path
    End If
    strAlertFile = strAlertFile & Application.PathSeparator & "ALERT_" & _
        Format(Now, "yyyy-mm-dd_hh:mm") & ".txt"
  End If
  
' Create error message from Err values
  Dim strAlert As String
  strAlert = "=========================================" & vbNewLine & _
    Now & " | " & strErrSource & vbNewLine & _
    lngErrNumber & ": " & strErrDescription
  
  
' if we can write a file there, write alert message
  If Utils.ParentDirExists(strAlertFile) = True Then
    Dim FileNum As Long
    FileNum = FreeFile()
    Open strAlertFile For Append As #FileNum
    Print #FileNum, strAlert
    Close #FileNum
  Else
    ' just in case it stays...
    DebugPrint strAlert
  End If
  
' Kill global objects
  If Not dictBookInfo Is Nothing Then
    Set dictBookInfo = Nothing
  End If
  If Not dictStyles Is Nothing Then
    Set dictStyles = Nothing
  End If
  If Not dictHeadings Is Nothing Then
    Set dictHeadings = Nothing
  End If
  If Not c_dictBaseStyle Is Nothing Then
    Set c_dictBaseStyle = Nothing
  End If
  If Not dictSections Is Nothing Then
    Set dictSections = Nothing
  End If
  If Not activeDoc Is Nothing Then
    Set activeDoc = Nothing
  End If

  ' Close all open documents
  Dim objDoc As Document
  Dim strExt As String
  For Each objDoc In Documents
    ' don't close any macro templates, might be running code.
    strExt = VBA.Right(objDoc.Name, InStr(StrReverse(objDoc.Name), "."))
    If strExt <> ".dotm" Then
      objDoc.Close wdDoNotSaveChanges
    End If
  Next objDoc

' Do NOT use `Exit Sub` before, since we ALWAYS want this to terminate.
ReportsTerminateError:
  End
End Sub


' ===== StyleCheck ============================================================
' Call this from origin project. Performs variety of style checks, returns
' dictionary containing results of various tests or whatever. Use private
' global variable to store the `dictStyles` object to access by later
' procedures.

Private Function StyleCheck(Optional FixUnstyled As Boolean = True) As _
  Dictionary

  On Error GoTo StyleCheckError
  
' At some point will also have to loop through active stories (EN. FN)
' Also `dictStyles` must be declared as global var.
  Set dictStyles = New Dictionary
  Dim dictReturn As Dictionary  ' the full dictionary object we'll return
  Dim dictInfo As Dictionary   ' sub-sub dict for indiv. style info
  
  Set dictReturn = New Dictionary
  dictReturn.Add "pass", False

' First test if our body style is even available in the doc (if not, not styled)
  If MacroHelpers.IsStyleInUse(strBodyStyle) = False Then
    dictReturn.Add "body_style_present", False
    Set StyleCheck = dictReturn
    Exit Function
  Else
    dictReturn.Add "body_style_present", True
    dictReturn.Add "unstyled_count", 0    ' for now, just a count. can add more data later
  End If

  Dim lngParaCt As Long: lngParaCt = activeDoc.Paragraphs.Count
'  DebugPrint "Total paragraphs: " & lngParaCt
  Dim rngPara As Range
  Dim strStyle As String
  Dim strBaseStyle As String
  Dim A As Long
  
' Loop through all paragraphs in document from END to START so we end up with
' FIRST page, and if we need to delete paras we don't mess up the count order
  For A = lngParaCt To 1 Step -1
  ' To break infinite loops.

  ' To do: increase? Add actual Err.Raise
    If A = 10000 Then
      DebugPrint "A = " & A
      Exit For
    End If
    
    If A Mod 200 = 0 Then
      DebugPrint "Paragraph " & A
    End If
    
  ' Set Range object for this paragraph
    Set rngPara = activeDoc.Paragraphs(A).Range
  ' Get style name
    strStyle = rngPara.ParagraphStyle


    
  ' If style name = Macmillan style...
    If Right(strStyle, 1) = ")" Then
    ' Is it a custom style? Assuming custom styles created correctly, should be
    ' based-on a Macmillan style we can revert to...
      strBaseStyle = activeDoc.Styles(strStyle).BaseStyle
      If Right(strBaseStyle, 1) = ")" Then
      ' ... though some good styles are based on other styles, so filter 'em

        If RevertToBaseStyle(strStyle) = True Then
          rngPara.Style = strBaseStyle
          strStyle = strBaseStyle
        End If
      End If
      
      
    ' If style does not exist in dict yet...
      If Not dictStyles.Exists(strStyle) Then
      ' ...create sub-dictionary
        Set dictInfo = New Dictionary
        dictInfo.Add "count", 0
        dictInfo.Add "start_paragraph", 0
        dictStyles.Add strStyle, dictInfo
        Set dictInfo = Nothing
      End If
    ' Increase style count and update start paragraph index
    ' .Item() method overwrites value for that key
      dictStyles(strStyle).Item("count") = dictStyles(strStyle)("count") + 1
      dictStyles(strStyle).Item("start_paragraph") = A
  ' Else (not Macmillan style)
    Else
    ' Increase unstyled count
      dictReturn.Item("unstyled_count") = dictReturn.Item("unstyled_count") + 1
      
    ' Change style, if requested
    ' To do: use logic to tag TX1, COTX1
    '        store style name externally
      If FixUnstyled = True Then
        rngPara.Style = strBodyStyle
      End If
    End If
  Next A
  
  ' What percentage are styled?
  Dim lngPercent As Single
  Dim blnPass As Boolean
  lngPercent = dictReturn("unstyledCount") / lngParaCt
  lngPercent = 1 - VBA.Round(lngPercent, 3)

' Threshold for "styled" is 50% of paragraphs have styles
  If lngPercent >= 0.5 Then
    blnPass = True
  Else
    blnPass = False
  End If
  
' update values in test dictionary
  dictReturn.Item("pass") = blnPass
  dictReturn.Item("unique_styles") = dictStyles.Count
  dictReturn.Item("percent_styled") = lngPercent
  
  Set StyleCheck = dictReturn
  Exit Function

StyleCheckError:
  Err.Source = strReports & "StyleCheck"
  If ErrorChecker(Err, strBodyStyle) = False Then
    Resume
  Else
    Call Reports.ReportsTerminate
  End If
End Function


' ===== IsbnCheck =============================================================
' Call this to run ISBN checks from the main Validator function. Optional param
' to determine if we should add the ISBN from the `book_info.json` file if not
' found.

Private Function IsbnCheck(Optional AddFromJson As Boolean = True) As _
  Dictionary
  On Error GoTo IsbnCheckError
   
 ' reset Error checker counter, so we can loop a few files
  lngErrorCount = 0
  Dim dictReturn As Dictionary
  Set dictReturn = New Dictionary
  dictReturn.Add "pass", False
  
' If no styled ISBN exists, try to find or add
  Dim blnStyledIsbn As Boolean
  blnStyledIsbn = FindIsbn(StyledOnly:=True)
  
  dictReturn.Add "styled_isbn", blnStyledIsbn
  
  If blnStyledIsbn = False Or AddFromJson = True Then
  
  ' Search for unstyled ISBN (if true, they are bookmarked)
    Dim blnUnstyled As Boolean
    blnUnstyled = FindIsbn()
    dictReturn.Add "unstyled_isbn", blnUnstyled
    
  ' convert bookmarks to styles
    Dim blnTagged As Boolean
    blnTagged = AddIsbnTags()
    dictReturn.Add "tag_unstyled_isbn", blnTagged
  End If

' Cleanup what ISBN tag is covering (should only be numerals, hyphens)
  Call Reports.ISBNcleanup
  
' Tag all URLs (to remove ISBN tag from ISBNs in URLs)
  Dim stStories(1 To 1) As WdStoryType
  stStories(1) = wdMainTextStory
  Call MacroHelpers.StyleAllHyperlinks(stStories)
  
' Read tagged isbns
  Dim isbnArray() As Variant
  isbnArray = MacroHelpers.GetText(strIsbnStyle, True)

' Add that this completed successfully
  If MacroHelpers.IsArrayEmpty(isbnArray) = False Then
    dictReturn.Item("pass") = True
    dictReturn.Add "list", isbnArray
  Else
    dictReturn.Item("pass") = False
  End If


' If we want to replace the given ISBNs with the one in book_info.json
  If AddFromJson = True Then
  ' Delete all tagged ISBNs
    Dim blnDeleteIsbns As Boolean
    blnDeleteIsbns = DeleteIsbns()
    dictReturn.Add "isbns_deleted", blnDeleteIsbns
    
  ' Add correct ISBN from JSON file
    Dim blnAddIsbn As Boolean
    blnAddIsbn = AddBookInfo(bk_ISBN)
    dictReturn.Add "isbn_added", blnAddIsbn
    
  ' Update `pass` key in test dictionary
    Dim blnPassed As Boolean
    blnPassed = MacroHelpers.IsStyleInUse(strIsbnStyle)
    dictReturn.Item("pass") = blnPassed
  End If

  Set IsbnCheck = dictReturn

  Exit Function
  
IsbnCheckError:
  Err.Source = strReports & "IsbnCheck"
  If ErrorChecker(Err, strIsbnStyle) = False Then
    Resume
  Else
    Call Reports.ReportsTerminate
  End If
End Function


' ===== FindIsbn ============================================================
' Searches for ISBNs (13-digits with or without hyphens). If found, tags with
' bookmarks and returns True. Default finds all ISBNs, StyledOnly:=True finds
' only those tagged with ISBN character style.

Private Function FindIsbn(Optional StyledOnly As Boolean = False) As Boolean
  On Error GoTo FindIsbnError
  Dim lngCounter As Long
  Dim strSearchPattern As String
  
  FindIsbn = False

' If looking for styled only, check that it's in use first (else not found)
  If StyledOnly = True Then
    If MacroHelpers.IsStyleInUse(strIsbnStyle) = False Then Exit Function
  End If
  
  activeDoc.Range.Select
  FindIsbn = False
  ' ISBN rules:
  ' * First 3 digits: 978 or 979
  ' * 4th digit: 0 or 1 (for English language)
  ' * next section: publisher/imprint, 2 to 7 digits
  ' * next section: book, 1 to 6 digits (these two total 8 digits)
  ' * last section: single check digit
  ' sections may or may not be separated by a hyphen, but note that you can't use
  ' {0,1} to search for "zero or one" occurrences of something.
  ' the below is OK for now. Try more specific if needed later.
  strSearchPattern = "97[89][0-9\-]{10,14}"
  
  ' lngCounter both to prevent infinite loop and also for array index
  ' which we want to start at 0 because may pass back to powershell
  lngCounter = -1
  
  MacroHelpers.zz_clearFind
  
  ' Start search at beginning of doc
  Selection.HomeKey Unit:=wdStory
  
  With Selection.Find
    .Text = strSearchPattern
    .Forward = True
    .Wrap = wdFindStop
    .MatchCase = True
    .MatchWildcards = True
    
    If StyledOnly = True Then
      .Format = True
      .Style = strIsbnStyle
    Else
      .Format = False
    End If

  End With

  ' If ISBNs are found, tag with a Bookmark
  Do While Selection.Find.Execute = True And lngCounter < 100
    FindIsbn = True
    lngCounter = lngCounter + 1

    ' Delete if bookmark already exists
    If activeDoc.Bookmarks.Exists("ISBN" & lngCounter) = True Then
      activeDoc.Bookmarks.Item("ISBN" & lngCounter).Delete
    End If
    ' Add bookmark for later procedures to pick up
    activeDoc.Bookmarks.Add "ISBN" & lngCounter, Selection
  Loop
  Exit Function

FindIsbnError:
  Err.Source = strReports & "FindIsbn"
  If ErrorChecker(Err, strIsbnStyle) = False Then
    Resume
  Else
    Call Reports.ReportsTerminate
  End If
End Function


' ===== DeleteIsbns ===========================================================
' Deletes all text with the ISBN style applied. Should expand at some point to
' use input style.

Private Function DeleteIsbns() As Boolean
  On Error GoTo DeleteIsbnsError
' Find all text with this style, replace with nothing
  MacroHelpers.zz_clearFind
  
  With activeDoc.Range.Find
    .Format = True
    .MatchWildcards = True
    .Style = strIsbnStyle
    .Text = "*"
    .Replacement.Text = vbNullString
    .Execute Replace:=wdReplaceAll
  End With
  
  Dim blnSuccess As Boolean
  blnSuccess = MacroHelpers.IsStyleInUse(strIsbnStyle)
  DeleteIsbns = Not blnSuccess

  Exit Function
  
DeleteIsbnsError:
  Err.Source = strReports & "DeleteIsbns"
  If ErrorChecker(Err, strIsbnStyle) = False Then
    Resume
  Else
    Call Reports.ReportsTerminate
  End If
End Function

' ===== AddBookInfo ===========================================================
' Add info from `book_info.json` to manuscript. Assume already know that it's
' not present. BookInfo is Private Enum.

Private Function AddBookInfo(InfoType As BookInfo) As Boolean
  On Error GoTo AddBookInfoError
  Dim strInfoKey As String
  Dim strInfoStyle As String ' style for new paragraph
  Dim enumParaDirection As WdCollapseDirection ' do we want to insert before or
                                           ' after found paragraph?
  Dim blnSearchFwd As Boolean ' Do we want to add to start of section?
  Dim lngStartPara As Long ' Default to first or last para?
  
' set vars based on what we're adding
  Select Case InfoType
    Case BookInfo.bk_Title
    ' Want to add BEFORE the FIRST Titlepage style
      strInfoKey = "title"
      strInfoStyle = strBookTitle
      enumParaDirection = wdCollapseStart
      blnSearchFwd = True
      lngStartPara = activeDoc.Paragraphs.Count
    Case BookInfo.bk_Authors
    ' Want to add AFTER the FIRST Titlepage style
    ' b/c Au not allowed section heading
      strInfoKey = "author"
      strInfoStyle = strAuthorName
      enumParaDirection = wdCollapseEnd
      blnSearchFwd = True
      lngStartPara = activeDoc.Paragraphs.Count
    Case BookInfo.bk_ISBN
    ' Want to add AFTER the LAST Copyright style
      strInfoKey = "isbn"
      strInfoStyle = strCopyright
      enumParaDirection = wdCollapseEnd
      blnSearchFwd = False
      lngStartPara = 1
  End Select

' ----- Find the section this should be added to ------------------------------
  Dim key1 As Variant
  Dim strInfoSection As String
  Dim colSectionStyles As Collection: Set colSectionStyles = New Collection
  strInfoSection = Left(strInfoStyle, InStr(strInfoStyle, " ") - 1)

' Find first paragraph in doc with style in same section as new text
' Styles were added to dictStyles in order they appear in the MS,
' So the first instance we find is the first present, ditto last
  For Each key1 In dictStyles.Keys
'    DebugPrint key1
    If InStr(key1, strInfoSection) > 0 Then
      colSectionStyles.Add key1
    End If
  Next key1
  
' Find index of first (or last) paragraph of each style in that section
  Dim lngCurrentStart As Long
  Dim varStyle As Variant
  Dim enumDocDirection As WdCollapseDirection

' if searching from end backward, collapse end
  If blnSearchFwd = True Then
    enumDocDirection = wdCollapseStart
  Else
    enumDocDirection = wdCollapseEnd
  End If
  
  If colSectionStyles.Count > 0 Then
    For Each varStyle In colSectionStyles
      MacroHelpers.zz_clearFind
      activeDoc.Select
      With Selection
        .Collapse enumDocDirection
        With .Find
          .Format = True
          .Style = varStyle
          .Forward = blnSearchFwd
          .Execute
        End With
      ' Get paragraph index of that paragraph
        lngCurrentStart = MacroHelpers.ParaIndex
      ' If we're looking for the LAST paragraph though, keep
      ' looping through the dictionary
        If blnSearchFwd = True Then
          If lngCurrentStart < lngStartPara Then
            lngStartPara = lngCurrentStart
          End If
        Else
          If lngCurrentStart > lngStartPara Then
            lngStartPara = lngCurrentStart
          End If
        End If
      
      End With
    Next varStyle
  End If

' Get string to add
  Dim strNewText As String
  If dictBookInfo.Exists(strInfoKey) = True Then
  ' If adding AFTER the LAST paragraph in the doc, add newline
  ' BEFORE the text, because we can't add after final paragraph mark
    If enumParaDirection = wdCollapseEnd And lngStartPara = _
      activeDoc.Paragraphs.Count Then
        strNewText = vbNewLine & dictBookInfo.Item(strInfoKey)
    Else
      strNewText = dictBookInfo.Item(strInfoKey) & vbNewLine
    End If
  Else
    AddBookInfo = False
    Exit Function
  End If
  
  ' Add text just before paragraph id'd above
  ' Once entered, new para takes index of lngStartPara.
  Dim rngNew As Range
  Set rngNew = activeDoc.Paragraphs(lngStartPara).Range
  With rngNew
    .Collapse enumParaDirection
    .InsertAfter strNewText
    .Style = strInfoStyle
  End With
  
  ' Test if it was successful
  If MacroHelpers.IsStyleInUse(strInfoStyle) = True Then
    AddBookInfo = True
  Else
    AddBookInfo = False
  End If
  
  ' ISBN also needs character style
  ' Search pattern copied from FindIsbn function. Maybe combine at some point.
  If InfoType = bk_ISBN Then
    MacroHelpers.zz_clearFind
    With rngNew.Find
      .MatchWildcards = True
      .Format = True
      .Text = "(97[89][0-9\-]{10,14})"
      .Replacement.Text = "\1"
      .Replacement.Style = strIsbnStyle
      .Execute Replace:=wdReplaceAll
    End With
  End If
  
  Exit Function
AddBookInfoError:
  Err.Source = strReports & "AddBookInfo"
  If ErrorChecker(Err, strInfoStyle) = False Then
    Resume
  Else
    Call Reports.ReportsTerminate
  End If
End Function


' ===== AddIsbnTags ===========================================================
' Converts bookmarked ISBNs to styles. FindIsbn function adds a bookmark
' that starts with "ISBN" in name to each.

Private Function AddIsbnTags() As Boolean
  On Error GoTo AddIsbnTagsError
  Dim bkName As Bookmark
  AddIsbnTags = False

  If MacroHelpers.IsStyleInDoc(strIsbnStyle) = False Then
    Dim myStyle As Style
    Set myStyle = activeDoc.Styles.Add(strIsbnStyle, wdStyleTypeCharacter)
  End If
  For Each bkName In activeDoc.Bookmarks
    If Left(bkName.Name, 4) = "ISBN" Then
      AddIsbnTags = True
      bkName.Select
      Selection.Style = strIsbnStyle
      bkName.Delete
    End If
  Next
  Exit Function
AddIsbnTagsError:
  Err.Source = strReports & "AddIsbnTags"
  If ErrorChecker(Err, strIsbnStyle) = False Then
    Resume
  Else
    Call Reports.ReportsTerminate
  End If
End Function


' ===== IsbnSearch ============================================================
' Implementation of IsbnCheck that returns a string, for Powershell call.

' Private because needs to be called independently from powershell if file name
' doesn't include ISBN. Optional FilePath is for passing doc path from PS.

' Don't actually need to log anything with LogFile param, but powershell expects
' to pass that argument so we'll make it optional.

'Private Function IsbnSearch(FilePath As String, Optional LogFile As String) _
'  As String
'  On Error GoTo IsbnSearchError
'
'' Make sure relevant file exists, is open
'  If Utils.IsOpen(FilePath) = False Then
'    Documents.Open FilePath
'  End If
'
'' Set reference to correct document
'  Set activeDoc = Documents(FilePath)
'
'' Create dictionary object to receive from IsbnCheck function
'  Dim dictIsbn As Dictionary
'  Set dictIsbn = New Dictionary
'  Set dictIsbn = IsbnCheck(AddFromJson:=False)
'
'' If ISBNs were found, they will be in the "list" element
'  If dictIsbn.Exists("list") = True Then
'  ' Reduce array elements to a comma-delimited string
'    IsbnSearch = Join(dictIsbn.Item("list"), ",")
'  Else
'    IsbnSearch = vbNullString
'  End If
'
'  activeDoc.Close wdDoNotSaveChanges
'  Set activeDoc = Nothing
'
'  Exit Function
'
'IsbnSearchError:
'  Err.Source = strReports & "IsbnSearch"
'  If ErrorChecker(Err, FilePath) = False Then
'    Resume
'  Else
'    Call Reports.ReportsTerminate
'  End If
'End Function


' ===== TitlepageCheck ========================================================
' Test that titlepage exists, Book Title exists, Author Name exists

Private Function TitlepageCheck() As Dictionary
  On Error GoTo TitlepageCheckError
' set up return info
  Dim dictReturn As Dictionary
  Set dictReturn = New Dictionary
  With dictReturn
    .Add "pass", False
    .Add "book_title_exists", False
    .Add "author_name_exists", False
  End With

  Dim blnTitle As Boolean
  Dim lngTitleCount As Long
  Dim strTitle As String
  Dim lngCount As Long: lngCount = 0
  Dim blnAuthor As Boolean

' Does Book Title exist?
  blnTitle = MacroHelpers.IsStyleInUse(strBookTitle)
  dictReturn.Item("book_title_exists") = blnTitle
  If blnTitle = False Then
    dictReturn.Item("book_title_added") = AddBookInfo(bk_Title)
  Else
    ' Is it more than one line?
    lngTitleCount = dictStyles(strBookTitle)("count")
    If lngTitleCount > 1 Then
      MacroHelpers.zz_clearFind
      activeDoc.Select
      Selection.HomeKey Unit:=wdStory
      With Selection
        .Find.Format = True
        .Find.Forward = True
        .Find.Style = strBookTitle
        .Find.Execute
        
        Do While .Find.Found = True And lngCount < 50
          lngCount = lngCount + 1
          .MoveDown Unit:=wdParagraph, Extend:=wdMove
          .Expand wdParagraph
          If .Style = strBookTitle Then
            .MoveUp Unit:=wdParagraph, Count:=2, Extend:=wdMove
            .Expand wdParagraph
            .Characters.Last.Delete
            .InsertAfter " "
          End If
          .Find.Execute
        Loop
      End With
    End If
  End If

' Does Author Name exist?
  blnAuthor = MacroHelpers.IsStyleInUse(strAuthorName)
  dictReturn.Item("author_name_exists") = blnAuthor
  If blnAuthor = False Then
    dictReturn.Item("author_name_added") = AddBookInfo(bk_Authors)
  End If

' Did it all work?
  If MacroHelpers.IsStyleInUse(strBookTitle) = True And _
    MacroHelpers.IsStyleInUse(strAuthorName) = True Then
    dictReturn.Item("pass") = True
  End If
  
  Set TitlepageCheck = dictReturn
  
  Exit Function
  
TitlepageCheckError:
  Err.Source = strReports & "TitlepageCheck"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call Reports.ReportsTerminate
  End If
End Function


' ===== SectionCheck ==========================================================
' Clean up some section styles, standardize page breaks, add section headings
' after each page break style.

Private Function SectionCheck() As Dictionary
  On Error GoTo SectionCheckError
  Dim dictReturn As Dictionary
  Set dictReturn = New Dictionary
  dictReturn.Add "pass", False

' Need to combine returned test dictionaries with this to reutrn single
  Dim dictStep As Dictionary
  
' fix some section styles (can probably cut after update style names)
  Set dictStep = StyleCleanup()
  Set dictReturn = ClassHelpers.MergeDictionary(dictReturn, dictStep)

' Fix page break formatting/styles
  Set dictStep = PageBreakCleanup()
  Set dictReturn = ClassHelpers.MergeDictionary(dictReturn, dictStep)
  
' Add section heads after each page break style
  Set dictStep = PageBreakCheck()
  Set dictReturn = ClassHelpers.MergeDictionary(dictReturn, dictStep)

  dictReturn.Item("pass") = True
  Set SectionCheck = dictReturn
  Exit Function

SectionCheckError:
  Err.Source = strReports & "SectionCheck"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call Reports.ReportsTerminate
  End If
End Function

' ===== StyleCleanup ==========================================================
' Tweaking some styles that will cause problems. Can probably cut once we
' update style name list.

Private Function StyleCleanup() As Dictionary
  On Error GoTo StyleCleanupError
  Dim dictReturn As Dictionary
  Set dictReturn = New Dictionary
  dictReturn.Add "pass", False
' Change "FM Epigraph" to just "Epigraph" so we can determine section
' Hard-code styles cuz we only need this until we change style names
  Dim strFmEpis(1 To 3) As String
  Dim X As Long
  strFmEpis(1) = strFmEpiText
  strFmEpis(2) = strFmEpiVerse
  strFmEpis(3) = strFmEpiSource

' Loop through to see if these are even in use. If yes, replace.
  Dim strNewStyle As String
  Dim blnSuccess As Boolean: blnSuccess = False
  For X = LBound(strFmEpis) To UBound(strFmEpis)
    If MacroHelpers.IsStyleInUse(strFmEpis(X)) = True Then
    ' Convert to correct style name (vbTextCompare = case insensitive)
      strNewStyle = VBA.LTrim(VBA.Replace(strFmEpis(X), "FM", "", _
        Compare:=vbTextCompare))
      blnSuccess = MacroHelpers.StyleReplace(strFmEpis(X), strNewStyle)
      dictReturn.Add "convert_fm_epigraph" & X, strNewStyle
    End If
  Next X

' Remove any section break characters. Can't assume they'll be in their own
' paragraphs, so add additional para break.
  MacroHelpers.zz_clearFind
  With activeDoc.Range.Find
    .Text = "^b"
    .Replacement.Text = "^p"
    .Execute Replace:=wdReplaceAll
  
    If .Found = True Then
      dictReturn.Add "delete_section_brk", True
    Else
      dictReturn.Add "delete_section_brk", False
    End If
  End With

' Convert any Section Break styles to Page Break (because they may have been
' added after a Page Break and we don't want to confuse PageBreak fixes later
' Use variable for new style name in case it's not present
  Dim strOldStyle As String
  strOldStyle = strSectionBreak
  strNewStyle = strPageBreak
  blnSuccess = MacroHelpers.StyleReplace(strOldStyle, strNewStyle)
  dictReturn.Add "delete_section_brk_style", blnSuccess

' Remove any Half Title paras. (If want to keep in future, create a separate
' function to search for all half titles, add headings/breaks.) Note that any
' extra page breaks will get cleaned up in `PageBreakCleanup` function.
  MacroHelpers.zz_clearFind
  strNewStyle = strHalftitle
  With activeDoc.Range.Find
    .Text = "*"
    .Replacement.Text = ""
    .Format = True
    .Style = strNewStyle
    .MatchWildcards = True
    .Execute Replace:=wdReplaceAll
  
    If .Found = True Then
      dictReturn.Add "delete_half_title", True
    Else
      dictReturn.Add "delete_half_title", False
    End If
  End With

  Call MacroHelpers.zz_clearFind
  
  dictReturn.Item("pass") = True
  Set StyleCleanup = dictReturn
  Exit Function

StyleCleanupError:
  Err.Source = strReports & "StyleCleanup"
  If ErrorChecker(Err, strNewStyle) = False Then
    Resume
  Else
    Call Reports.ReportsTerminate
  End If
End Function

' ===== IsHeading =============================================================
' Is this paragraph style a Macmillan heading style? Eventually store style
' names externally.

Private Function IsHeading(StyleName As String) As Boolean
  On Error GoTo IsHeadingError
' `dictHeadings` is global scope so only have to create once
  If dictHeadings Is Nothing Then
    Set dictHeadings = New Dictionary
  ' Check for `headings.json` file, read into global dictionary
    Dim strHeadings As String
    strHeadings = Environ("BkmkrScripts") & Application.PathSeparator & _
      "Word-template_assets" & Application.PathSeparator & "headings.json"
    If Utils.IsItThere(strHeadings) = True Then
      Set dictHeadings = ClassHelpers.ReadJson(strHeadings)
    Else
      Err.Raise MacError.err_FileNotThere
    End If
  End If
  
' So just see if our style is one of these styles
  Dim blnResult As Boolean
  blnResult = dictHeadings.Exists(StyleName)
  If blnResult = True Then
    dictHeadings.Item(StyleName) = True
  End If
  IsHeading = blnResult
  Exit Function

IsHeadingError:
  Err.Source = strReports & "IsHeading"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call Reports.ReportsTerminate
  End If
End Function


' ===== RevertToBaseStyle =====================================================
' Is this paragraph style a Macmillan heading style? Eventually store style
' names externally.

Private Function RevertToBaseStyle(StyleName As String) As Boolean
  On Error GoTo RevertToBaseStyleError
' Hard code for now. `c_dictBaseStyle` is global scope so only have to create once
  If c_dictBaseStyle Is Nothing Then
    Set c_dictBaseStyle = New Dictionary
  ' Value arg is the style key is based on, tho we don't need it yet
    With c_dictBaseStyle
      .Add "Titlepage Logo (logo)", "Bookmaker Processing Instruction (bpi)"
      .Add "About Author Text Head (atah)", "BOB Ad Title (bobt)"
      .Add "BM Title (bmt)", "BM Head (bmh)"
      .Add "bookmaker loosen (bkl)", "bookmaker keep together (kt)"
      .Add "Bookmaker Page Break (br)", "Page Break (pb)"
      .Add "Bookmaker Processing Instruction (bpi)", "Design Note (dn)"
      .Add "bookmaker tighten (bkt)", "bookmaker keep together (kt)"
      .Add "Chap Title Nonprinting (ctnp)", "Chap Number (cn)"
      .Add "Column Break (cbr)", "Column Head (ch)"
      .Add "Extract - Bullet List (extbl)", "Extract-No Indent (ext1)"
      .Add "Extract - Diary (extd)", "Extract (ext)"
      .Add "Extract - Inscription (ins)", "Extract (ext)"
      .Add "Extract - Newspaper (news)", "Extract (ext)"
      .Add "Extract - Num List (extnl)", "Extract - Bullet List (extbl)"
      .Add "Extract - Telegram (tel)", "Extract - Email (extem)"
      .Add "Extract - Transcript (trans)", "Extract (ext)"
      .Add "Extract Source (exts)", "Extract - Website (extws)"
      .Add "FM Head ALT (afmh)", "FM Head (fmh)"
      .Add "FM Text ALT (afmtx)", "FM Text (fmtx)"
      .Add "FM Text No-Indent ALT (afmtx1)", "FM Text (fmtx)"
      .Add "FM Title (fmt)", "FM Head (fmh)"
      .Add "Part Epigraph - non-verse (pepi)", "Part Epigraph - verse (pepiv)"
      .Add "Part Opening Text (potx)", "Text - Standard (tx)"
      .Add "Part Opening Text No-Indent (potx1)", "Text - Std No-Indent (tx1)"
      .Add "Teaser Opening Text No-Indent (totx1)", "Teaser Opening Text (totx)"
      .Add "TOC Author (cau)", "TOC Frontmatter Head (cfmh)"
      .Add "TOC Page Number (cnum)", "TOC Backmatter Head (cbmh)"
    End With
  End If
  
' So if our style is here, we DON'T want to revert, so we reverse it
  Dim blnResult As Boolean
  blnResult = c_dictBaseStyle.Exists(StyleName)
'  DebugPrint blnResult
  RevertToBaseStyle = Not blnResult
  
  Exit Function

RevertToBaseStyleError:
  Err.Source = strReports & "RevertToBaseStyle"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call Reports.ReportsTerminate
  End If
End Function


' ===== SectionName ===========================================================
' Determines section heading text and style from original para's style name.
' Reads from external JSON file. When we get around to changing style names so
' the section is always the first word we can make this much simpler.

' JSON Format:
'    "Bibliography": {
'      "text":"Bibliography",
'      "headingStyle":"BM Head Nonprinting (bmhnp)"
'      }

' Object names for each need to be added to (1) SectionsJson enumeration and (2)
' first select statement below

Private Function SectionName(StyleName As String, Optional JsonString As _
  SectionsJson = j_text) As String
  On Error GoTo SectionNameError
  Dim dictItem As Dictionary

' Create dictionary from JSON if it hasn't been created yet
  If dictSections Is Nothing Then
  ' Check for `sections.json` file, read into global dictionary
    Dim strSections As String
    strSections = Environ("BkmkrScripts") & Application.PathSeparator & _
      "Word-template_assets" & Application.PathSeparator & "sections.json"
    If Utils.IsItThere(strSections) = True Then
      Set dictSections = ClassHelpers.ReadJson(strSections)
    Else
      Err.Raise MacError.err_FileNotThere
    End If
  End If

' JSON key = first word in style passed to us
  Dim strFirst As String
  strFirst = Left(StyleName, InStr(StyleName, " ") - 1)
  ' DebugPrint strFirst
' If style is in JSON...
  If dictSections.Exists(strFirst) = True Then
  ' ... get object for that style.
    Set dictItem = dictSections.Item(strFirst)
' Else, just make it a generic chapter heading
  Else
    Set dictItem = dictSections.Item("Chap")
  End If
  
' Convert enum to string. Default is "text"
  Dim strJsonString As String

  Select Case JsonString
    Case j_text
      strJsonString = "text"
    Case j_style
      strJsonString = "headingStyle"
  End Select

' Retrieve value
  If dictItem.Exists(strJsonString) Then
    SectionName = dictItem.Item(strJsonString)
    ' DebugPrint SectionName
  End If
  Exit Function

SectionNameError:
  Err.Source = strReports & "SectionName"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    ReportsTerminate
  End If
End Function


' ===== AddHeading ============================================================
' Adds *NP heading ABOVE the paragraph passed as arg, with section name text.
' Though each chapter is just "Chapter" -- will have to add numbers later.

Private Function AddHeading(paraInd As Long) As Boolean
  On Error GoTo AddHeadingError
' Get current para count to test at end that we added a new one
  Dim lngParas As Long
  lngParas = activeDoc.Paragraphs.Count
  
' Set range for the para in question
  Dim rngPara As Range
  Set rngPara = activeDoc.Paragraphs(paraInd).Range
  
' Get style name of that paragraph
  Dim strParaStyle As String
  strParaStyle = rngPara.ParagraphStyle

' Use style name to get text and style for heading
  Dim strSectionName As String
  Dim strHeadingStyle As String
  strSectionName = SectionName(strParaStyle, j_text)
  strHeadingStyle = SectionName(strParaStyle, j_style)

' add line ending ('cuz new paragraph), insert as new paragraph
  strSectionName = strSectionName & vbNewLine
  rngPara.InsertBefore (strSectionName)
  
' Add correct style (inserted paragraph now part of `rngPara` object)
' ErrorChecker will add style if it doesn't exist
'  Debug.Print strHeadingStyle
  rngPara.Paragraphs(1).Style = strHeadingStyle

' Verify we added a paragraph
  Dim lngNewParas As Long
  lngNewParas = activeDoc.Paragraphs.Count
  If lngNewParas = lngParas + 1 Then
    AddHeading = True
  Else
    AddHeading = False
  End If
  Exit Function
  
AddHeadingError:
  Err.Source = strReports & "AddHeading"
  If ErrorChecker(Err, strHeadingStyle) = False Then
    Resume
  Else
    Call Reports.ReportsTerminate
  End If
End Function

' ===== PageBreakCleanup ======================================================
' Clean up page break characters/styles, so just single paragraph break chars
' styles as "Page Break" remain.

Private Function PageBreakCleanup() As Dictionary
  On Error GoTo PageBreakCleanupError
  Dim dictReturn As Dictionary
  Set dictReturn = New Dictionary
  dictReturn.Add "pass", False
  
' Add paragraph breaks around every page break character (so we know for sure
' paragraph style of break won't apply to any body text). Will add extra blank
' paragraphs that we can clean up later.
' Also add "Page Break (pb)" style.

  MacroHelpers.zz_clearFind
  With activeDoc.Range.Find
    .Text = "^m"
    .Replacement.Text = "^p^m^p"
    .Format = True
    .Replacement.Style = strPageBreak
    .Execute Replace:=wdReplaceAll
    
    If .Found = True Then
      dictReturn.Add "page_brk_replace", True
    Else
      dictReturn.Add "page_brk_replace", False
    End If
  End With

' If we had an unstyled page break char, new trailing ^p is wrong style
' Use this to make sure all correct style.
  MacroHelpers.zz_clearFind
  With activeDoc.Range.Find
    .Text = "^m^13{1,}"
    .Replacement.Text = "^m^p"
    .Format = True
    .MatchWildcards = True
    .Replacement.Style = strPageBreak
    .Execute Replace:=wdReplaceAll
    
    If .Found = True Then
      dictReturn.Add "page_brk_replace2", True
    Else
      dictReturn.Add "page_brk_replace2", False
    End If
  End With

' Now that we are sure every PB char has PB style, remove all PB char
  MacroHelpers.zz_clearFind
  With activeDoc.Range.Find
    .Text = "^m"
    .Replacement.Text = ""
    .Execute Replace:=wdReplaceAll
    
    If .Found = True Then
      dictReturn.Add "page_brk_remove", True
    Else
      dictReturn.Add "page_brk_remove", False
    End If
  End With

' Remove multiple PB-styled paragraphs in a row
  MacroHelpers.zz_clearFind
  With activeDoc.Range.Find
    .Text = "^13{2,}"
    .Replacement.Text = "^p"
    .Format = True
    .Style = strPageBreak
    .MatchWildcards = True
    .Execute Replace:=wdReplaceAll
    
    If .Found = True Then
      dictReturn.Add "rm_multiple_paras", True
    Else
      dictReturn.Add "rm_multiple_paras", False
    End If
  End With

' Remove any first paragraphs until we get one with text
  Dim lngCount As Long
  Dim rngPara1 As Range
  Dim strKey As String
  
  Do
    lngCount = lngCount + 1
    strKey = "firstParaPB" & lngCount
    Set rngPara1 = activeDoc.Paragraphs.First.Range
    If MacroHelpers.IsNewLine(rngPara1.Text) = True Then
      dictReturn.Add strKey, True
      rngPara1.Delete
    Else
      dictReturn.Add strKey, False
      Exit Do
    End If
  Loop Until lngCount > 20 ' For runaway loops
  
  dictReturn.Item("pass") = True
  Set PageBreakCleanup = dictReturn

  Exit Function

PageBreakCleanupError:
  Err.Source = strReports & "PageBreakCleanup"
  If ErrorChecker(Err, strPageBreak) = False Then
    Resume
  Else
    Call Reports.ReportsTerminate
  End If
End Function

' ===== PageBreakCheck ========================================================
' Check that every page break is followed by a heading. If not, add one.

Private Function PageBreakCheck() As Dictionary
  On Error GoTo PageBreakCheckError
  Dim dictReturn As Dictionary
  Set dictReturn = New Dictionary
  dictReturn.Add "pass", False

' Loop through search of all "Page Break"-styled paragraphs
  Dim strNextStyle As String
  Dim lngParaInd As Long
  
  Dim lngParaCount As Long
  lngParaCount = activeDoc.Paragraphs.Count
'  DebugPrint "Total paragraphs: " & lngParaCount
  Dim lngCount As Long
  lngCount = 0
  
  MacroHelpers.zz_clearFind
  Selection.HomeKey Unit:=wdStory
  With Selection.Find
    .Text = ""
    .Style = strPageBreak
    .Execute
    
    Do While .Found = True And lngCount < 200
      ' Loop counter
      lngCount = lngCount + 1
      lngParaInd = MacroHelpers.ParaIndex
      ' DebugPrint "Page break: " & lngParaInd
      ' Errors if we try to access para after end, so check that
      If lngParaCount > lngParaInd Then
      ' If the NEXT paragraph is NOT an approved heading style...
        strNextStyle = activeDoc.Paragraphs(lngParaInd + 1).Range.ParagraphStyle
        ' DebugPrint "Next para style: " & strNextStyle
        If IsHeading(strNextStyle) = False Then
          ' DebugPrint "Next style is NOT heading"
          ' ... add a CTNP heading
          If AddHeading(lngParaInd + 1) = True Then
'            DebugPrint "Heading added"
          ' We added a paragraph, so increase count in whole doc
            lngParaCount = lngParaCount + 1
            dictReturn.Item("added_headings") = dictReturn.Item("added_headings") + 1
          End If
        End If
      End If
      .Execute
    Loop
  End With
  
' Remove all "Page Break" styles. We need to add a page break to the end of
' each section later (in case some have no page break), but we don't want to
' duplicate so we'll remove now. Do not remove the actual page break, because
' we need a non-heading style between a section that is ONLY a heading (maybe
' it's a placeholder) and the next section's heading.

  If MacroHelpers.StyleReplace(strPageBreak, strBodyStyle) = True Then
    dictReturn.Add "pg_brk_style_removed", True
  Else
    dictReturn.Add "pg_brk_style_removed", False
  End If

  dictReturn.Item("pass") = True
  Set PageBreakCheck = dictReturn
  Exit Function
  
PageBreakCheckError:
  Err.Source = strReports & "PageBreakCheck"
  If ErrorChecker(Err, strPageBreak) = False Then
    Resume
  Else
    Call Reports.ReportsTerminate
  End If
End Function

' ===== SectionHeadInd ========================================================
' Identify the start of each section (based on paragraph style). Returns an
' array of paragraph indices.

Private Function SectionHeadInd() As Variant
  On Error GoTo SectionHeadIndError
  Dim strCurrentStyle As String
  Dim blnIsHeading As Boolean
  Dim paraInd() As Variant
  Dim lngParaCount As Long: lngParaCount = activeDoc.Paragraphs.Count
' P = current paragraph, Q = paraInd() upper-bound
  Dim P As Long, Q As Long
 
' Loop through paragraphs with DO loop, because we area going to be skipping
' ahead later in the loop (i.e., don't want to start loop with EVERY para
  P = 1
  Q = 0
  Do Until P > lngParaCount
    strCurrentStyle = activeDoc.Paragraphs(P).Range.ParagraphStyle
'    DebugPrint strCurrentStyle
    If IsHeading(strCurrentStyle) = True Then
    ' This is the FIRST heading paragraph in a row, add to output array
      Q = Q + 1
      ReDim Preserve paraInd(1 To Q)
      paraInd(Q) = P
'      DebugPrint "Heading index: " & P
    ' Loop until we find the next paragraph that is NOT a heading (assumes that
    ' allowable heading sections are all grouped together. Would get confused
    ' if someone throws in a non-heading style between headings!
      Do
        P = P + 1
        If P < lngParaCount Then
          strCurrentStyle = activeDoc.Paragraphs(P).Range.ParagraphStyle
          blnIsHeading = IsHeading(strCurrentStyle)
        Else
          Exit Do
        End If
      Loop Until blnIsHeading = False
    End If
    P = P + 1
  Loop
  
  SectionHeadInd = paraInd()
  Exit Function

SectionHeadIndError:
  Err.Source = strReports & "SectionHeadInd"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call Reports.ReportsTerminate
  End If
End Function


' ===== SectionRanges =========================================================
' Create an array of range objects, one for each section of the manuscript.
' Pass the array generated by SectionHeadInd() as an argument.

Private Function SectionRange(ParaIndexArray() As Variant) As Variant
  On Error GoTo SectionRangeError
  Dim rangeArray() As Variant
  Dim rngSection As Range
  Dim lngParaCount As Long: lngParaCount = activeDoc.Paragraphs.Count
  
  Dim lngLBound As Long
  Dim lngUBound As Long
  lngLBound = LBound(ParaIndexArray)
  lngUBound = UBound(ParaIndexArray)
  ReDim Preserve rangeArray(lngLBound To lngUBound)
  
' G is array index number
  Dim g As Long
  Dim lngStart As Long
  Dim lngEnd As Long

' Loop through passed array
  For g = lngLBound To lngUBound
  ' Determine start and end section index numbers
    lngStart = ParaIndexArray(g)
'    DebugPrint lngStart
    If g < lngUBound Then
      lngEnd = ParaIndexArray(g + 1) - 1
    Else
      lngEnd = lngParaCount
    End If
'    DebugPrint lngEnd
    Dim lngColor As Long
  ' Set range based on those start/end points
    With activeDoc
      Set rngSection = .Range(Start:=.Paragraphs(lngStart).Range.Start, _
        End:=.Paragraphs(lngEnd).Range.End)
      
'      ' DEBUGGING
'      If G Mod 2 = 0 Then
'        lngColor = wdColorAqua
'      Else
'        lngColor = wdColorPink
'      End If
'      rngSection.Shading.BackgroundPatternColor = lngColor
'    ' DEBUGGING
    
    
    End With
  ' Add range to array
    Set rangeArray(g) = rngSection
  Next g
  
  SectionRange = rangeArray()
  Exit Function
  
SectionRangeError:
  Err.Source = strReports & "SectionRange"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call Reports.ReportsTerminate
  End If
End Function


' ===== HeadingCheck ==========================================================
' Validate a variety of heading requirements, fix if not met. Parameter is an
' array of ranges (one for each section) returned from SectionRange function

Private Function HeadingCheck() As Dictionary
  On Error GoTo HeadingCheckError
  Dim dictReturn As Dictionary
  Set dictReturn = New Dictionary
  dictReturn.Add "pass", False

' Verify first paragraph is a heading
  Dim dictStep As Dictionary
  Set dictStep = Reports.FirstParaCheck
  Set dictReturn = ClassHelpers.MergeDictionary(dictReturn, dictStep)

' Get paragraph indices of section start paragraphs
  Dim rngParaInd() As Variant
'  DebugPrint "SectionHeadInd start"
  rngParaInd = SectionHeadInd
'  DebugPrint "SectionHeadInd stop"

' Create array of ranges of each section
  Dim rngSections() As Variant
'  DebugPrint "SectionRange start"
  rngSections = SectionRange(rngParaInd)
'  DebugPrint "SectionRange end"
  
' Loop through section ranges
  Dim D As Long
  Dim rngSect As Range
  Dim rngFirst As Range
  Dim rngSecond As Range
  Dim strFirstStyle As String
  Dim strSecondStyle As String
  Dim strFirstText As String
  Dim strSecondText As String
  Dim lngChapCount As Long: lngChapCount = 0
  Dim strSectionKey As String
  Dim blnRmCharFormat As Boolean
  Dim blnChapterNumber As Boolean: blnChapterNumber = False
  Dim rngPara1 As Range
  Dim rngParaLast As Range

  For D = UBound(rngSections) To LBound(rngSections) Step -1
    ' Replace with error message for infinite loop
    If D Mod 10 = 0 Then
       DebugPrint "Checking section " & D
    End If

    strSectionKey = "section" & D
'    DebugPrint strSectionKey
    
    Set rngSect = rngSections(D)
    
'    '' DEBUGGING
'    Dim lngColor As Long
'    If D Mod 2 = 0 Then
'      lngColor = wdColorGold
'    Else
'      lngColor = wdColorGreen
'    End If
'
'    rngSect.Shading.BackgroundPatternColor = lngColor
'    '' DEBUGGING
  
  ' Get style name and text (w/o new line) of first paragraph in range.
    Set rngFirst = rngSect.Paragraphs(1).Range
    strFirstStyle = rngFirst.ParagraphStyle
    strFirstText = Left(rngFirst.Text, Len(rngFirst.Text) - 1)
    
  ' Test if range is more than 1 paragraph
    Dim blnMultiPara As Boolean
    If rngSect.Paragraphs.Count >= 2 Then
      blnMultiPara = True
    Else
      blnMultiPara = False
    End If
    
  ' If section only has 1 paragraph (e.g., backmatter section TK), trying to
  ' get 2nd will result in an error. If more than 1, get style name and text
  ' of 2nd paragraph.
    If blnMultiPara = True Then
      Set rngSecond = rngSect.Paragraphs(2).Range
      strSecondStyle = rngSecond.ParagraphStyle
      strSecondText = Left(rngSecond.Text, Len(rngSecond.Text) - 1)
    Else
      strSecondStyle = vbNullString
      strSecondText = vbNullString
    End If

'    DebugPrint "strFirstStyle: " & strFirstStyle
'    DebugPrint "strFirstText: " & strFirstText
'    DebugPrint "strSecondStyle: " & strSecondStyle
'    DebugPrint "strSecondText: " & strSecondText
'
    dictReturn.Add strSectionKey & "FirstStyle", strFirstStyle
    Select Case strFirstStyle
      Case strChapNonprinting, c_strFmHeadNonprinting, c_strBmHeadNonprinting
      ' Does it have text?
        If strFirstText = vbNullString Then  ' No text (just paragraph return)
        ' Add section name to blank paragraph
          strFirstText = SectionName(strFirstStyle)
          rngFirst.InsertBefore strFirstText
          dictReturn.Add strSectionKey & "AddHeadText", strFirstText
        Else
        ' Is that text just "Chapter" with no numbers?
          If strFirstText = "Chapter" Then
            blnChapterNumber = True
          End If
        End If
        
      Case strChapNumber
      ' Remove any character styles
        blnRmCharFormat = ChapNumCleanUp(StyleName:=strChapNumber, SearchRange:=rngFirst)
        dictReturn.Add strSectionKey & "ChapNumCleanUp", blnRmCharFormat
        
      ' Is next para a CT? (If no, change this to CT)
        If strSecondStyle <> strChapTitle Then
          rngFirst.Style = strChapTitle
          dictReturn.Add strSectionKey & "ChapNumToTitle", True
        End If
        
      Case strPartNumber
      ' Remove any character styles
        blnRmCharFormat = ChapNumCleanUp(StyleName:=strPartNumber, SearchRange:=rngFirst)
        dictReturn.Add strSectionKey & "PartNumCleanUp", blnRmCharFormat
        
      ' Is next para a PT? (If no, change this to CT)
        If strSecondStyle <> strPartTitle Then
          rngFirst.Style = strPartTitle
          dictReturn.Add strSectionKey & "PartNumToTitle", True
        End If

      Case strFmHead
      ' Remove any character styles
        blnRmCharFormat = ChapNumCleanUp(StyleName:=strFmHead, SearchRange:=rngFirst)
        dictReturn.Add strSectionKey & "FmhCleanUp", blnRmCharFormat
      
      Case strBmHead
      ' Remove any character styles
        blnRmCharFormat = ChapNumCleanUp(StyleName:=strBmHead, SearchRange:=rngFirst)
        dictReturn.Add strSectionKey & "BmhCleanUp", blnRmCharFormat

      Case strChapTitle
      ' Is next para CN?
        If strSecondStyle = strChapNumber Then
        ' Remove any character styles
          blnRmCharFormat = ChapNumCleanUp(StyleName:=strChapNumber, SearchRange:=rngSecond)
          dictReturn.Add strSectionKey & "ChapNumCleanUp", blnRmCharFormat
        ' move CN before CT
          rngSecond.Cut
          rngFirst.Collapse (wdCollapseStart)
          rngFirst.PasteAndFormat (wdFormatOriginalFormatting)
          dictReturn.Add strSectionKey & "ChapTitleSwap", True
        End If
        
      ' Is next para ALSO a CT?
        If strSecondStyle = strChapTitle Then
        ' Combine into single para (delete paragraph return)
          rngFirst.Characters.Last.Select
          Selection.Delete
        End If
  
      Case strPartTitle
        ' Is next para PN?
        If strSecondStyle = strPartNumber Then
        ' Remove any character styles
          blnRmCharFormat = ChapNumCleanUp(StyleName:=strPartNumber, SearchRange:=rngSecond)
          dictReturn.Add strSectionKey & "PartNumCleanUp", blnRmCharFormat
        ' move PN before PT
          rngSecond.Cut
          rngFirst.Collapse (wdCollapseStart)
          rngFirst.PasteAndFormat (wdFormatOriginalFormatting)
          dictReturn.Add strSectionKey & "PartTitleSwap", True
        End If
      ' Is next para ALSO a PT?
        If strSecondStyle = strPartTitle Then
        ' Combine into single para (delete paragraph return)
          rngFirst.Characters.Last.Select
          Selection.Delete
        End If

      Case strFmTitle
      ' Is next para FMH?
        If strSecondStyle = strFmHead Then
        ' Remove any character styles
          blnRmCharFormat = ChapNumCleanUp(StyleName:=strFmHead, SearchRange:=rngSecond)
          dictReturn.Add strSectionKey & "FmhCleanUp", blnRmCharFormat
        ' move PN before PT
          rngSecond.Cut
          rngFirst.Collapse (wdCollapseStart)
          rngFirst.PasteAndFormat (wdFormatOriginalFormatting)
          dictReturn.Add strSectionKey & "FmTitleSwap", True
        Else
          rngFirst.Style = strFmHead
          dictReturn.Add strSectionKey & "FmtToFmh", True
        End If
      
      Case strBmTitle
      ' Is next para BMH?
        If strSecondStyle = strBmHead Then
        ' Remove any character styles
          blnRmCharFormat = ChapNumCleanUp(StyleName:=strBmHead, SearchRange:=rngSecond)
          dictReturn.Add strSectionKey & "BmhCleanUp", blnRmCharFormat
        ' move PN before PT
          rngSecond.Cut
          rngFirst.Collapse (wdCollapseStart)
          rngFirst.PasteAndFormat (wdFormatOriginalFormatting)
          dictReturn.Add strSectionKey & "BmTitleSwap", True
        Else
          rngFirst.Style = strBmHead
          dictReturn.Add strSectionKey & "BmtToBmh", True
        End If
    End Select
    
  ' Add section breaks to START of range, i.e. end of section before
  ' this one. Don't add to first range, though.
  ' Get separate Range objects for first and last paragraphs, because
  ' collapse method changes the range. Also, if have a single-paragraph
  ' range, still need to add section before and PB after.
    Set rngPara1 = rngSect.Paragraphs(1).Range
    If D > LBound(rngSections) Then
      rngPara1.Collapse Direction:=wdCollapseStart
      rngPara1.InsertBreak Type:=wdSectionBreakNextPage
      dictReturn.Add strSectionKey & "AddSectionBreak", True
    End If
    
    Set rngParaLast = rngSect.Paragraphs.Last.Range
    If D < UBound(rngSections) Then
       rngParaLast.Collapse Direction:=wdCollapseEnd
       rngParaLast.InsertAfter vbNewLine
       rngParaLast.Style = Reports.strPageBreak
       dictReturn.Add strSectionKey & "AddPageBreak", True
    End If
  Next D
  
' Reset Note Options to restart numbering at each section?
  dictReturn.Item("pass") = True
  Set HeadingCheck = dictReturn
  Exit Function
  
HeadingCheckError:
  Err.Source = strReports & "HeadingCheck"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call Reports.ReportsTerminate
  End If
End Function

' ===== FirstParaCheck ========================================================
' Make sure first paragraph is an acceptable heading style, or add new para.

Private Function FirstParaCheck() As Dictionary
  On Error GoTo FirstParaCheckError
    Dim dictReturn As Dictionary
    Set dictReturn = New Dictionary
    dictReturn.Add "pass", False
    
    Dim rngPara1 As Range
    Dim strStyle As String
    Dim strHeadingText As String
    Dim strHeadingStyle As String
    Set rngPara1 = activeDoc.Paragraphs.First.Range
    strStyle = rngPara1.ParagraphStyle
    
    If IsHeading(strStyle) = False Then
      dictReturn.Add "headingStyle", False
      strHeadingText = "Frontmatter" & vbNewLine
      strHeadingStyle = c_strFmHeadNonprinting
      rngPara1.InsertBefore (strHeadingText)
      activeDoc.Paragraphs.First.Style = strHeadingStyle
      dictReturn.Add "fmHeadAdded", True
      dictReturn.Item("pass") = True
    Else
      dictReturn.Add "headingStyle", True
      dictReturn.Item("pass") = True
    End If
    
    Set FirstParaCheck = dictReturn
  Exit Function
  
FirstParaCheckError:
  Err.Source = strReports & "FirstParaCheck"
  If ErrorChecker(Err, strHeadingStyle) = False Then
    Resume
  Else
    Call Reports.ReportsTerminate
  End If
End Function

' ===== IllustrationCheck =====================================================
' Various illustration validation checks.

Private Function IllustrationCheck() As Dictionary
  On Error GoTo IllustrationCheckError
  Dim dictReturn As Dictionary
  Set dictReturn = New Dictionary
  dictReturn.Add "pass", False
  
' replace all "Illustration holder" text with placeholder.
' eventually make this its own function?
  Dim strPlaceholder As String
  strPlaceholder = "TK.jpg" & vbNewLine

  If MacroHelpers.IsStyleInUse(strIllustrationHolder) = False Then
    dictReturn.Add "ill_holder_exists", False
  Else
    dictReturn.Add "ill_holder_exists", True
  ' Replace text of all paragraphs with that style
    MacroHelpers.zz_clearFind
    With activeDoc.Range.Find
      .Text = ""
      .Replacement.Text = "^p"
      .Format = True
      .Style = strIllustrationHolder
      .Execute Replace:=wdReplaceAll
' Two steps: remove text (except para return), then add new text
' Don't call clearFind, want to maintain style info.
      .Text = "^p"
      .Replacement.Text = "tk.jpg^p"
      .Execute Replace:=wdReplaceAll
      
      If .Found = True Then
        dictReturn.Add "replace_ill_text", True
      Else
        dictReturn.Add "replace_ill_text", False
      End If
    
    End With
    MacroHelpers.zz_clearFind
  End If
  
  dictReturn.Item("pass") = True
  Set IllustrationCheck = dictReturn
  Exit Function
  
IllustrationCheckError:
  Err.Source = strReports & "IllustrationCheck"
  If ErrorChecker(Err) = False Then
    Resume
  Else
    Call Reports.ReportsTerminate
  End If
End Function


'
Private Sub ISBNcleanup()
'removes "span ISBN (isbn)" style from all but the actual ISBN numerals

    'check if that style exists, if not then exit sub
    On Error GoTo ErrHandler:
        Dim keyStyle As Word.Style
        Set keyStyle = activeDoc.Styles("span ISBN (isbn)")

    Dim strISBNtextArray()
    ReDim strISBNtextArray(1 To 3)

    strISBNtextArray(1) = "-[!0-9]"     'any hyphen followed by any non-digit character
    strISBNtextArray(2) = "[!0-9]-"     'any hyphen preceded by any non-digit character
    strISBNtextArray(3) = "[!-0-9]"     'any character other than a hyphen or digit

    ' re: above--need to search for hyphens first, because if you lead with what is now 3, you
    ' remove the style from any characters around hyphens, so if you search for a hyphen next to
    ' a character later, it won't return anything because the whole string needs to have the
    ' style applied for it to be found.

    Dim g As Long
    For g = LBound(strISBNtextArray()) To UBound(strISBNtextArray())

        'Move selection to start of document
        Selection.HomeKey Unit:=wdStory

        With Selection.Find
            .ClearFormatting
            .Text = strISBNtextArray(g)
            .Replacement.ClearFormatting
            .Replacement.Text = ""
            .Forward = True
            .Wrap = wdFindStop
            .Format = True
            .Style = "span ISBN (isbn)"                     'find this style
            .Replacement.Style = "Default Paragraph Font"   'replace with this style
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = True
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With

        Selection.Find.Execute Replace:=wdReplaceAll

    Next g

Exit Sub

ErrHandler:
  'Style doesn't exist in document
    If Err.Number = 5941 Or Err.Number = 5834 Then
      Exit Sub
    Else
      Err.Source = strReports & "ISBNcleanup"
      If MacroHelpers.ErrorChecker(Err) = False Then
        Resume
      Else
        Call Reports.ReportsTerminate
      End If
    End If

End Sub

'
Private Function ChapNumCleanUp(Optional StyleName As String = strChapNumber, _
  Optional SearchRange As Range) As Boolean
  On Error GoTo ChapNumCleanUpError
  ' Removes character styles from Chapter Number paragraphs
  ChapNumCleanUp = False
  
' Set selection
  If SearchRange Is Nothing Then
    activeDoc.Select
  Else
    SearchRange.Select
'    DebugPrint SearchRange.Paragraphs.Count
  End If

' Move selection back to start of RANGE
  Selection.Collapse wdCollapseStart
  MacroHelpers.zz_clearFind

  Dim intCount As Long
  intCount = 0
  With Selection.Find
    .Forward = True
    .Format = True
    .Style = StyleName
    .Execute
  ' < 1000 to prevent infinite loop
'    Do While .Found = True And intCount < 1000
''      ChapNumCleanUp
'      intCount = intCount + 1
      #If Mac Then
      ' Mac 2011 doesn't support ClearCharacterFormattingAll method
      ' And ClearFormatting removes paragraph formatting as well
        Selection.ClearFormatting
        Selection.Style = StyleName
      #Else
        Selection.ClearCharacterDirectFormatting
        Selection.Style = "Default Paragraph Font"
'        Selection.ClearCharacterAllFormatting
      #End If
'      Selection.Collapse wdCollapseEnd
'    Loop
  End With

  Exit Function

ChapNumCleanUpError:
  Err.Source = strReports & "ChapNumCleanUp"
  If ErrorChecker(Err, StyleName) = False Then
    Resume
  Else
    Reports.ReportsTerminate
  End If
End Function


