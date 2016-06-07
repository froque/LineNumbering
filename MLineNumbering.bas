Attribute VB_Name = "MLineNumbering"
Option Explicit

Private Declare Function GetShortPathName Lib "kernel32" Alias _
"GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Const csEXEHELPTEXT As String = "Usage: LNT /Pproject /Odirectory  [/C[directory]] [/W] [/Lincrement]" & vbCr & vbCr _
                            & "/P - " & vbTab & " Project to generate line numbers" & vbCr _
                            & "/O - " & vbTab & " Output directory for new source code (default of \LN)" & vbCr _
                            & "/C - " & vbTab & " Compile the project with line numbers and place out in specified directory" & vbCr _
                            & "/W - " & vbTab & " Wipe the Output directory before starting" & vbCr _
                            & "/T - " & vbTab & " Conditional Compilation Arguments" & vbCr _
                            & "/M - " & vbTab & " Maintain the same Path32 (build path) in the new line numbered project " & vbCr _
                            & "/I - " & vbTab & " Line increment to use (default of 1)" & vbCr _
                            & "/VMajor.Minor.Revision(y/n) - " & vbTab & " New version number (auto increment)" & vbCr _
                            & "/? - " & vbTab & " Display this help text"

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Private gsProject           As String
Private gsCompileDir        As String
Private gsOutputDir         As String
Private gbClearOutputDir    As Boolean
Private gbCompileProject    As Boolean
Private glIncrement         As Long
Private gbMaintainPaths     As Boolean
Private gsOriginalOutDir    As String
Private gbChangeVersion     As Boolean
Private gsMajor             As String
Private gsMinor             As String
Private gsRevision          As String
Private gbAutoIncrement     As Boolean
Private gsConditionalArgs   As String

Private Const LINE_CONTINUATION = "_"
Private Const GSSQ       As String = """"

Private Const END_SUB = "End Sub"
Private Const END_FUNCTION = "End Function"
Private Const END_PROPERTY = "End Property"

Private Const SUB_LINE = "Sub "
Private Const PUBLIC_SUB = "Public Sub "
Private Const Private_SUB = "Private Sub "
Private Const FRIEND_SUB = "Friend Sub "

Private Const FUNCTION_LINE = "Function "
Private Const PUBLIC_FUNCTION = "Public Function "
Private Const PRIVATE_FUNCTION = "Private Function "
Private Const FRIEND_FUNCTION = "Friend Function "

Private Const PROPERTY_LINE = "Property "
Private Const PUBLIC_PROPERTY = "Public Property "
Private Const PRIVATE_PROPERTY = "Private Property "
Private Const FRIEND_PROPERTY = "Friend Property "

Private Const MODULE_LINE = "Module="
Private Const CLASS_LINE = "Class="
Private Const USERCONTROL_LINE = "UserControl="
Private Const FORM_LINE = "Form="
Private Const RELATEDDOC_LINE = "RelatedDoc="
Private Const RESFILE_LINE = "ResFile32="
Private Const COMPATIBLEEXE32_LINE = "CompatibleEXE32="
Private Const PATH32_LINE = "Path32="
Private Const MAJORVER_LINE = "MajorVer="
Private Const MINORVER_LINE = "MinorVer="
Private Const REVISIONVER_LINE = "RevisionVer="
Private Const AUTOINCREMENTVER_LINE = "AutoIncrementVer="

Private Const CASE_STATEMENT = "Case"
Private Const SELECTCASE_STATEMENT = "Select Case"
Private Const ENDSELECT_STATEMENT = "End Select"
Private Const CONDITIONAL_STATEMENT = "#"
Private Const COMMENT_STATEMENT = "'"
Private Const REM_STATEMENT = "Rem"

Public Sub Main()
    On Error GoTo errTrap

    'set defaults
    gsProject = ""
    gsOutputDir = "\LN"
    glIncrement = 1
    gbClearOutputDir = False
    
    Dim bCancel As Boolean
    If fnbParseCommandLine(bCancel) = False Then
        MsgBox App.Title & vbCr & vbCr & csEXEHELPTEXT, vbInformation, App.Title & " Help"
        Call ExitWithErrorLevel(1)
        Exit Sub
    ElseIf bCancel = True Then
        Call ExitWithErrorLevel(0)
        Exit Sub
    End If
    
    'OK We have what we need from the Command Line
        
    'Clean up the Output Directory
    If Right$(gsOutputDir, 1) = "\" Then
        gsOutputDir = Left$(gsOutputDir, Len(gsOutputDir) - 1)
    End If
        
    gsOriginalOutDir = gsOutputDir
        
    If Left$(gsOutputDir, 1) = "\" Then
        Dim sProjectDir As String
        Dim sProjectFileName As String
        gsProject = GetLongFilename(gsProject)
        GetProjectFileNameAndDir gsProject, sProjectFileName, sProjectDir
        gsOutputDir = sProjectDir & gsOutputDir
    End If
        
    If Len(Dir(gsOutputDir, vbDirectory)) > 0 Then
        If gbClearOutputDir = False Then
            If Len(Dir(gsOutputDir & "\*.*")) > 0 Then
                If vbYes = MsgBox("Files exist in the directory " & gsOutputDir & vbCr & vbCr & "Do you wish to clear this directory?", vbYesNo) Then
                    ClearDirectory gsOutputDir
                End If
            End If
        Else
            ClearDirectory gsOutputDir
        End If
    Else
        MkDir gsOutputDir
    End If
                     
    If fnbParseProjectFile(gsProject, gsOutputDir) = True Then
        If gbCompileProject Then
            If fnbCompileProject(sProjectDir, gsOutputDir, sProjectFileName) Then
            Else
                Call ExitWithErrorLevel(1)
                Exit Sub
            End If
        End If
    Else
        Call ExitWithErrorLevel(1)
        Exit Sub
    End If
    
   
    'Exit with success
    MsgBox "Completed Successfully.", vbInformation, "Line Numbering Tool"
    Call ExitWithErrorLevel(0)
    Exit Sub
    
errTrap:
    MsgBox "Main Error: " & Err.Description & IIf(Erl, ", Line:" & Erl, "")

End Sub

Private Sub ClearDirectory(ByRef sPath As String)
    On Error GoTo errTrap
    
    Dim sFile As String
    Dim sDir As String
  
    If Len(sPath) = 0 Then Exit Sub
  
    If Right$(sPath, 1) = "\" Then
        sDir = sPath
    Else
        sDir = sPath & "\"
    End If
  
    sFile = Dir(sDir & "*.*")
  
    Do While Len(sFile) > 0
                          
        'Make sure the file isn't read only
        SetAttr sDir & sFile, vbNormal
        
        'Bye...
        Kill sDir & sFile
        
        sFile = Dir
    Loop
  
    Exit Sub
    
errTrap:
200       MsgBox "ClearDirectory Error: " & Err.Description & IIf(Erl, ", Line:" & Erl, "")
End Sub
Private Function fnbParseCommandLine(ByRef Cancel As Boolean) As Boolean
    On Error GoTo errTrap
    
    'Parse the Command Line
    Dim vCmds   As Variant
    Dim v       As Variant
    Dim sCmds   As String
    Dim sCmd    As String
    Dim sValue  As String
    Dim bShowHelp As Boolean
    
    sCmds = Command$
    
    If Len(sCmds) = 0 And InIDE Then
        sCmds = "/PLineNumbering.vbp " '/C"
    End If
    
    'sCmds = Replace(sCmds, "-", "/", 1, 1)  'Replace a leading "-" if necessary
    'sCmds = Replace(sCmds, " -", " /") 'Replace all remaining "-"s with "/"
    sCmds = " " & sCmds
    
    vCmds = Split(sCmds, " /")

    For Each v In vCmds
        If Len(v) > 0 Then
            sCmd = Left(v, 1)
            sValue = ""
            If (Len(v) > 1) Then
                sValue = Trim(Right(v, Len(v) - 1))
            End If
            Select Case UCase(sCmd)
                Case "P" 'project
                    If Len(sValue) > 0 Then
                        If Left$(sValue, 1) = """" And Right$(sValue, 1) = """" Then
                            sValue = Mid$(sValue, 2, Len(sValue) - 2)
                        End If
                        If Left$(sValue, 1) = "\" Then
                            sValue = Mid$(sValue, 2)
                        End If
                        gsProject = sValue
                    End If
                Case "O" 'output directory
                    If Len(sValue) > 0 Then
                        If Left$(sValue, 1) = """" Then
                            sValue = Mid$(sValue, 2, Len(sValue) - 1)
                        End If
                        If Right$(sValue, 1) = """" Then
                            sValue = Mid$(sValue, 1, Len(sValue) - 1)
                        End If
                        If Left$(sValue, 1) <> "\" Then
                            sValue = "\" & sValue
                        End If
                        gsOutputDir = sValue
                    End If
                Case "C" 'Compile directory
                    gbCompileProject = True
                    If Len(sValue) > 0 Then
                        If Left$(sValue, 1) = """" And Right$(sValue, 1) = """" Then
                            sValue = Mid$(sValue, 2, Len(sValue) - 2)
                        End If
                        gsCompileDir = sValue
                    End If
                Case "T" 'Conditional Compilation
                    gsConditionalArgs = sValue
                Case "W" 'Wipe the output directory
                    gbClearOutputDir = True
                Case "M" 'Maintain build path
                    gbMaintainPaths = True
                Case "I" 'increment
                    glIncrement = Val(sValue)
                    If glIncrement < 1 Then glIncrement = 1
                    If glIncrement > 1000 Then glIncrement = 1000
                Case "?" 'Help Text
                    bShowHelp = True
                Case "H"
                    If UCase$(sValue) = "ELP" Then bShowHelp = True
                Case "V"
                    If Len(sValue) > 0 Then
                        Dim sValueArr() As String
                        sValueArr = Split(sValue, ".")
                        gsMajor = 0
                        gsMinor = 0
                        gsRevision = 0
                        If UBound(sValueArr) >= 2 Then
                            gsMajor = sValueArr(0)
                            gsMinor = sValueArr(1)
                            gsRevision = Val(sValueArr(2))
                            gbAutoIncrement = IIf(LCase(Right(sValueArr(2), 1)) = "y", True, False)
                            gbChangeVersion = True
                        End If
                    End If
            End Select
        End If
    Next

    If gsProject = "" Then bShowHelp = True
    
    'Show the Help screen if we couldn't parse correctly
    If bShowHelp = False Then
      fnbParseCommandLine = True
    End If
    
    Exit Function
    
errTrap:
    MsgBox "fnbParseCommandLine Error: " & Err.Description & IIf(Erl, ", Line:" & Erl, "")

End Function

Private Function fnbCompileProject(ByRef sOldProjectDir As String, ByRef sNewProjectDir As String, ByRef sProjectFile As String) As Boolean
    On Error GoTo errTrap
    
    Dim sCmd As String
    Dim sTempDir As String
    Dim lCount As Long
    Dim sProject As String
    Dim iFileNumber As Integer
    Dim sFile As String
    
    Const VBEXE As String = "C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"
    
    If Dir(VBEXE) = "" Then
      MsgBox "Unable to compile, can not find VB program" & vbCr & vbCr & VBEXE, vbCritical
      Exit Function
    End If
    
    lCount = 0
    sTempDir = sOldProjectDir & "\Temp"
    
    Do While Len(Dir(sTempDir, vbDirectory)) > 0
        lCount = lCount + 1
        sTempDir = sOldProjectDir & "\Temp" & lCount
    Loop
    
    'Now we have a temporary sub directory
    If fnbMoveDirectoryFiles(sOldProjectDir, sTempDir) = False Then
        MsgBox "The moving of the project to a temporary directory failed.  Please check the status of the project, as it may be missing files", vbCritical
        Exit Function
    End If
    
    'Now lets move the line numbered source into the original directory
    If fnbMoveDirectoryFiles(sNewProjectDir, sOldProjectDir) = False Then
        MsgBox "The moving of the line numbered project to the original directory failed.  Please check the status of the project, as it may be missing files", vbCritical
        Exit Function
    End If
    
    'Build the Command line to execute to build the project
    sProject = sOldProjectDir & "\" & sProjectFile
    
    sCmd = GSSQ & VBEXE & GSSQ & " /m " & GSSQ & sProject & GSSQ & " /out " & GSSQ & sOldProjectDir & "\Build.Log" & GSSQ
    
    If Len(gsCompileDir) > 0 Then
        If Len(Dir(gsCompileDir, vbDirectory)) = 0 Then
            MkDir gsCompileDir
        End If
        sCmd = sCmd & " /outdir " & GSSQ & gsCompileDir & GSSQ
    End If
    
    'sCmd = sCmd & " /d   " 'Remove the conditional compilation constants
    sCmd = sCmd & " /d" & gsConditionalArgs
    
    'Remove the old log file (if it exists)
    KillFile sOldProjectDir & "\Build.Log"
    
    'Now do the compile
    ShellAndWaitToFinish sCmd
    
    'Update Path32 in line numbered project
    If gbMaintainPaths Then
        fnbCorrectPath32InProjectFile sProjectFile, sOldProjectDir, sNewProjectDir
    End If
    
    'Now move everything back to where it should be...
    If fnbMoveDirectoryFiles(sOldProjectDir, sNewProjectDir) = False Then
        MsgBox "The moving of the compiled line numbered project to the sub directory failed.  Please check the status of the project, as it may be missing files", vbCritical
        Exit Function
    End If
    
    If fnbMoveDirectoryFiles(sTempDir, sOldProjectDir) = False Then
        MsgBox "The moving of the original project to the original directory failed.  Please check the status of the project, as it may be missing files", vbCritical
        Exit Function
    End If
        
    RmDir sTempDir
    
    'Log file should have been created
    If Len(Dir(sNewProjectDir & "\Build.Log")) > 0 Then
        iFileNumber = FreeFile
        
        Open sNewProjectDir & "\Build.Log" For Input As #iFileNumber
        sFile = Input(LOF(iFileNumber), iFileNumber)
        Close iFileNumber
    
        If InStr(1, sFile, " succeeded.") = 0 Then
            sFile = "The " & sProjectFile & " project has not been compiled." & vbCr & vbCr _
                    & Trim$(sFile)
            
            Call MsgBox(sFile, vbCritical)
        
            Exit Function
        End If
    End If
    
    fnbCompileProject = True
    
    Exit Function
    
errTrap:
    MsgBox "fnbCompileProject Error: " & Err.Description & IIf(Erl, ", Line:" & Erl, "")

End Function
Private Function fnbParseProjectFile(ByRef sProject As String, ByRef sOutputDir As String) As Boolean
    On Error GoTo errTrap

    Dim iInputFileNumber As Integer
    Dim iOutputFileNumber As Integer
    Dim iOrigOutputFileNumber As Integer
    Dim sOriginalLine As String
    Dim sLine As String
    Dim sOutput As String
    Dim sProjectDir As String
    Dim sProjectFileName As String
    Dim sOriginalProjectFile As String
    Dim bGetFile As Boolean
    Dim bCopyFile As Boolean
    Dim bParseFile As Boolean
    Dim bCopyRenameFile As Boolean
    Dim sFile As String
    Dim sFileName As String
    Dim iFilePos As Integer
    Dim bCheckFRX As Boolean
    Dim bCheckCTX As Boolean
    Dim bAutoInc As Boolean
    Dim sTemp As String
    
    GetProjectFileNameAndDir sProject, sProjectFileName, sProjectDir
        
    'Open a new source file for writing
    iOutputFileNumber = FreeFile
        
    Open sOutputDir & "\" & Trim(sProjectFileName) For Output As iOutputFileNumber
    
    'Open the source file for reading
    iInputFileNumber = FreeFile
    Open sProjectDir & "\" & sProjectFileName For Input As iInputFileNumber
        
    If gbCompileProject Then
        sOriginalProjectFile = sProjectDir & "\TMP_" & Trim(sProjectFileName)
        'We need to increment the Revision Number of the original Source
        KillFile sOriginalProjectFile
        iOrigOutputFileNumber = FreeFile
        Open sOriginalProjectFile For Output As iOrigOutputFileNumber
    End If
    
    Do Until EOF(iInputFileNumber)
        'Get a line from the file
        Line Input #iInputFileNumber, sOriginalLine
        
        'Trim any spaces from the beginning
        sLine = Trim(sOriginalLine)
        sOutput = sLine
        
        'Don't test an empty line
        If Len(sLine) > 0 Then
        
            bGetFile = False                'We need to process the file
            bParseFile = False              'We need to add line numbers
            bCopyFile = False               'Copy File only
            bCopyRenameFile = False         'Copy and rename (used for compatible)
            bCheckFRX = False
            bCheckCTX = False
            
            'Check for Forms,Modules,UserControls or Classes
            If InStr(1, sLine, MODULE_LINE) = 1 Or _
                    InStr(1, sLine, CLASS_LINE) = 1 Then
                bGetFile = True
                bParseFile = True
            ElseIf InStr(1, sLine, USERCONTROL_LINE) = 1 Then
                bGetFile = True
                bParseFile = True
                bCheckCTX = True
            ElseIf InStr(1, sLine, FORM_LINE) = 1 Then
                bGetFile = True
                bParseFile = True
                bCheckFRX = True
            'Now check for Related Documents and Resource File
            ElseIf InStr(1, sLine, RELATEDDOC_LINE) = 1 Or _
                    InStr(1, sLine, RESFILE_LINE) = 1 Then
                bGetFile = True
                bCopyFile = True
            'Now adjust the CompatibleExe if required...
            ElseIf InStr(1, sLine, COMPATIBLEEXE32_LINE) = 1 Then
                bGetFile = True
                bCopyRenameFile = True
            ElseIf InStr(1, sLine, PATH32_LINE) = 1 Then
                sOutput = PATH32_LINE & Chr$(34) & Chr$(34)
                If gbMaintainPaths Then
                    sOutput = sLine
                End If
            ElseIf InStr(1, sLine, MAJORVER_LINE) = 1 Then
                iFilePos = InStr(1, sLine, "=")
                If gbChangeVersion Then
                    sOriginalLine = MAJORVER_LINE & gsMajor
                    sOutput = MAJORVER_LINE & gsMajor
                End If
            ElseIf InStr(1, sLine, MINORVER_LINE) = 1 Then
                iFilePos = InStr(1, sLine, "=")
                If gbChangeVersion Then
                    sOriginalLine = MINORVER_LINE & gsMinor
                    sOutput = MINORVER_LINE & gsMinor
                End If
            ElseIf InStr(1, sLine, REVISIONVER_LINE) = 1 Then
                iFilePos = InStr(1, sLine, "=")
                If gbChangeVersion Then
                    If gbAutoIncrement Then
                        sOriginalLine = REVISIONVER_LINE & (Val(gsRevision) + 1)
                    Else
                        sOriginalLine = REVISIONVER_LINE & gsRevision
                    End If
                    sOutput = REVISIONVER_LINE & gsRevision
                Else
                    sTemp = Trim$(Mid$(sLine, iFilePos + 1))
                    If IsNumeric(sTemp) Then
                        sOriginalLine = REVISIONVER_LINE & CStr(Val(sTemp) + 1)
                    End If
                End If
            ElseIf InStr(1, sLine, AUTOINCREMENTVER_LINE) = 1 Then
                If gbChangeVersion Then
                    sOutput = AUTOINCREMENTVER_LINE & IIf(gbAutoIncrement, "1", "0")
                    sOriginalLine = AUTOINCREMENTVER_LINE & IIf(gbAutoIncrement, "1", "0")
                Else
                    iFilePos = InStr(1, sLine, "=")
                    sTemp = Trim$(Mid$(sLine, iFilePos + 1))
                    If IsNumeric(sTemp) Then
                        If Val(sTemp) = 1 Then bAutoInc = True
                    End If
                End If
                
            End If
                            
            If bGetFile Then
               
                'Is the line of the format (Module=SrchGlobals; SrchGlobals.bas)
                iFilePos = InStr(1, sLine, "; ")
                If iFilePos <= 0 Then
                    'Is the line of the format (Form=SummaryFrm.frm)
                    iFilePos = InStr(1, sLine, "=")
                    If iFilePos >= 0 Then
                        'Step past the "="
                        iFilePos = iFilePos + 1
                    End If
                Else
                    'Step past the "; "
                    iFilePos = iFilePos + 2
                End If
               
                'After all that did we get a file name?
                If iFilePos > 0 Then
                    sFile = Mid(sLine, iFilePos)
                    
                    'Trim the quotes from either side, if they exist.
                    If Left$(sFile, 1) = """" And Right$(sFile, 1) = """" Then
                        sFile = Mid$(sFile, 2, Len(sFile) - 2)
                    End If
                    
                    
                    'Get just the File Name
                    If InStr(sFile, "\") > 0 Then
                        sFileName = Right$(sFile, Len(sFile) - InStrRev(sFile, "\"))
                        sFile = sProjectDir & "\" & sFile
                    Else
                        sFileName = sFile
                        sFile = sProjectDir & "\" & sFileName
                    End If
                    
                    sOutput = Left(sLine, iFilePos - 1) & sFileName
                    If Len(sFile) Then
                        If bParseFile Then
                            'For code add the numbers
                            If AddLineNumbers(sFile, sOutputDir) = False Then
                                MsgBox "Unable to add line numbers to File :" & vbCr & vbCr & sFile, vbCritical
                                Exit Function
                            End If
                        ElseIf bCopyFile Then
                            'Just copt the file over
                            If Dir(sFile) = "" Then
                                MsgBox "Unable to find File :" & vbCr & vbCr & sFile, vbCritical
                                Exit Function
                            End If
                            FileCopy sFile, sOutputDir & "\" & sFileName
                        ElseIf bCopyRenameFile Then
                            'Copy File over and rename it (to avoid conflicts)
                            If Left$(sFile, 2) = ".." Then
                              sFile = sProjectDir & "\" & sFile
                            End If
                            
                            If Dir(sFile) = "" Then
                                MsgBox "Unable to find File :" & vbCr & vbCr & sFile, vbCritical
                                Exit Function
                            End If
                            
                            sOutput = Left(sLine, iFilePos - 1) & "CMP_" & sFileName
                            FileCopy sFile, sOutputDir & "\" & "CMP_" & sFileName
                        End If
                        
                        If bCheckFRX Then
                            'If the file is a form, check for an FRX
                            If UCase$(Right$(sFile, 3)) = "FRM" Then
                                sFile = Left$(sFile, Len(sFile) - 3) & "frx"
                                sFileName = Left$(sFileName, Len(sFileName) - 3) & "frx"
                                
                                If Len(Dir(sFile)) > 0 Then
                                    FileCopy sFile, sOutputDir & "\" & sFileName
                                End If
                            End If
                        End If
                    
                        If bCheckCTX Then
                            'If the file is a user control, check for an CTX
                            If UCase$(Right$(sFile, 3)) = "CTL" Then
                                sFile = Left$(sFile, Len(sFile) - 3) & "ctx"
                                sFileName = Left$(sFileName, Len(sFileName) - 3) & "ctx"
                                
                                If Len(Dir(sFile)) > 0 Then
                                    FileCopy sFile, sOutputDir & "\" & sFileName
                                End If
                            End If
                        End If
                    End If
                
                End If
            End If
        End If
        
        'Output the line
        Print #iOutputFileNumber, sOutput
        
        If gbCompileProject Then
            Print #iOrigOutputFileNumber, sOriginalLine
        End If
    Loop
    
    Close iInputFileNumber
    Close iOutputFileNumber
    Close iOrigOutputFileNumber

    If gbCompileProject Then
        If bAutoInc Or gbChangeVersion Then
            'We need to replace the original project file with the updated one (with new revision level)
            If Len(Dir(sOriginalProjectFile)) > 0 And Len(Dir(sProjectDir & "\" & sProjectFileName)) > 0 Then
                KillFile sProjectDir & "\" & sProjectFileName
                Name sOriginalProjectFile As sProjectDir & "\" & sProjectFileName
            End If
        End If
    End If

    On Error Resume Next
    Kill sOriginalProjectFile
    
    fnbParseProjectFile = True
    
    Exit Function
    
errTrap:
    MsgBox "fnbParseProjectFile Error: " & Err.Description & IIf(Erl, ", Line:" & Erl, "")
End Function

Private Sub GetProjectFileNameAndDir(sProject As String, sProjectFileName As String, _
    sProjectDir As String)
            
    'Retrive the Project File
    If InStr(sProject, "\") > 0 Then
        sProjectFileName = Right$(sProject, Len(sProject) - InStrRev(sProject, "\"))
        sProjectDir = Left$(sProject, Len(sProject) - Len(sProjectFileName) - 1)
    Else
        sProjectFileName = sProject
        sProjectDir = App.Path
    End If
    
End Sub

Private Function AddLineNumbers(ByRef sFile As String, ByRef sOutputDir As String) As Boolean
    On Error GoTo errTrap
    
    Dim sFileDir As String
    Dim sFileName As String
    
    Dim iInputFileNumber As Integer
    Dim iOutputFileNumber As Integer
    Dim sLine As String
    Dim sTrimmedLine As String
    Dim bSkipNextLine As Boolean
    Dim bSkipThisLine As Boolean
    Dim bInProc As Boolean
    Dim bStartOfSelect As Boolean
    Dim iLineNumberCount As Integer
    Dim sLineNumberStr As String * 8  'MAX 5 Characters (allows 99,999,999 lines per module!)
    Dim sFirstToken As String
    Dim bFoundNumbers As Boolean
    
    'Retrieve the File Name
    
    Dim sOldFileName As String
    sOldFileName = sFile
        
    GetProjectFileNameAndDir sFile, sFileName, sFileDir
    
'    If InStr(sFile, "\") > 0 Then
'        sFileName = Right$(sFile, Len(sFile) - InStrRev(sFile, "\"))
'        sFileDir = Left$(sFile, Len(sFile) - Len(sFileName) - 1)
'    Else
'        sFileName = sFile
'        sFileDir = App.Path
'    End If
    
    'Open a new source file for writing
    iOutputFileNumber = FreeFile
    Open sOutputDir & "\" & sFileName For Output As #iOutputFileNumber
    
    'Open the source file for reading
    iInputFileNumber = FreeFile
    Open sFileDir & "\" & sFileName For Input As #iInputFileNumber
    
    'Reset the flags
    bInProc = False
    iLineNumberCount = glIncrement
    
    'Loop through the file
    Do While Not EOF(iInputFileNumber)
    
        Line Input #iInputFileNumber, sLine
        
        sTrimmedLine = Trim(sLine)
    
        'What do we have left?
        If Len(sTrimmedLine) = 0 Then
        
            'Don't add comments to blank lines
            bSkipThisLine = True
        
        Else
            'Handle flags for this line
            If bSkipNextLine = True Then
                'No numbers to be added for this line
                bSkipThisLine = True
                
                'Reset the flag
                bSkipNextLine = False
            Else
                'As far as we know, process this line
                bSkipThisLine = False
            End If
            
            'Don't check continuation lines
            If Right$(sTrimmedLine, 1) = LINE_CONTINUATION Then
                bSkipNextLine = True
            End If
            
            'Are we leaving a procedure?
            If bInProc Then
                If InStr(1, sTrimmedLine, END_SUB) = 1 Or _
                        InStr(1, sTrimmedLine, END_FUNCTION) = 1 Or _
                        InStr(1, sTrimmedLine, END_PROPERTY) = 1 Then
                    'Outside Proc
                    bInProc = False
                    bSkipThisLine = True
                End If
                
                If InStr(1, sTrimmedLine, SELECTCASE_STATEMENT) = 1 Then
                    bStartOfSelect = True 'Used to track if in between a select and the first case statement
                End If
                If InStr(1, sTrimmedLine, ENDSELECT_STATEMENT) = 1 Then
                    bStartOfSelect = False 'Used to track if in between a select and the first case statement
                End If
                
                If InStr(1, sTrimmedLine, CASE_STATEMENT) = 1 Then
                    bStartOfSelect = False 'Used to track if in between a select and the first case statement
                    
                    bSkipThisLine = True
                    
                    'Add in spaces to keep justified...
                    If bFoundNumbers = False Then
                      sLine = Space$(Len(sLineNumberStr)) & sLine
                    End If
                End If
                
                If InStr(1, sTrimmedLine, CONDITIONAL_STATEMENT) = 1 Then
                    bSkipThisLine = True
                    
                    'Add in spaces to keep justified...
                    If bFoundNumbers = False Then
                      sLine = Space$(Len(sLineNumberStr)) & sLine
                    End If
                End If

                If InStr(1, sTrimmedLine, REM_STATEMENT) = 1 _
                    Or InStr(1, sTrimmedLine, COMMENT_STATEMENT) Then
                    
                    'Not allowed to put line numbers between a select and the first case statement
                    If bStartOfSelect = True Then
                        bSkipThisLine = True
                        
                        'Add in spaces to keep justified...
                        If bFoundNumbers = False Then
                          sLine = Space$(Len(sLineNumberStr)) & sLine
                        End If
                    End If
                End If

                'Don't worry about these types of lines, compiles OK, but we keep the code justified...
                'DIM_STATEMENT, STATIC_STATEMENT

                
            'Are we entering a procedure?
            Else
                If InStr(1, sTrimmedLine, SUB_LINE) = 1 Or InStr(1, sTrimmedLine, PUBLIC_SUB) = 1 _
                        Or InStr(1, sTrimmedLine, Private_SUB) = 1 Or InStr(1, sTrimmedLine, FRIEND_SUB) = 1 _
                        Or InStr(1, sTrimmedLine, FUNCTION_LINE) = 1 Or InStr(1, sTrimmedLine, PUBLIC_FUNCTION) = 1 _
                        Or InStr(1, sTrimmedLine, PRIVATE_FUNCTION) = 1 Or InStr(1, sTrimmedLine, FRIEND_FUNCTION) = 1 _
                        Or InStr(1, sTrimmedLine, PROPERTY_LINE) = 1 Or InStr(1, sTrimmedLine, PUBLIC_PROPERTY) = 1 _
                        Or InStr(1, sTrimmedLine, PRIVATE_PROPERTY) = 1 Or InStr(1, sTrimmedLine, FRIEND_PROPERTY) = 1 Then
                    'Inside Proc
                    bInProc = True
                    bSkipThisLine = True
                End If
            End If
        
        End If
        
        'Are we in a procedure and not skipping this line
        If bInProc = True And bSkipThisLine = False Then
            'Add a line number
            sLineNumberStr = CStr(iLineNumberCount)
            
            'Check if any existing numbers
            sFirstToken = Split(sTrimmedLine, " ")(0)
            
            If IsNumeric(sFirstToken) Then
                bFoundNumbers = True
            
                sLine = sTrimmedLine
                
                Dim lNumberChars    As Long
                Dim lExtraChars     As Long
                Dim lPosToken       As Long
                
                lNumberChars = Len(sFirstToken)
                lExtraChars = Len(sLineNumberStr) - lNumberChars
                lPosToken = InStr(sLine, sFirstToken)
                
                'Trimming the Number off the line, and also trying to keep everything justified...
                If lNumberChars <= Len(sLineNumberStr) Then
                    sLine = Mid$(sLine, lPosToken + lNumberChars)
                    If Left$(sLine, lExtraChars) = Space$(lExtraChars) Then
                        sLine = Mid$(sLine, lExtraChars + 1)
                    End If
                Else
                    sLine = Mid$(sLine, lPosToken + lNumberChars)
                End If
            End If
            
            sLine = sLineNumberStr & sLine
            Print #iOutputFileNumber, sLine
            iLineNumberCount = iLineNumberCount + glIncrement
        Else
            Print #iOutputFileNumber, sLine
        End If
    Loop
    
    Close iInputFileNumber
    Close iOutputFileNumber
    
    AddLineNumbers = True
    
    Exit Function
    
errTrap:
    MsgBox "AddLineNumbers Error: " & Err.Description & IIf(Erl, ", Line:" & Erl, "")
End Function

'Private Function InIDE() As Boolean
'    On Error GoTo errTrap
'
'    Debug.Print 1 / 0
'
'    InIDE = False
'
'    Exit Function
'
'errTrap:
'    InIDE = True
'End Function

Private Function InIDE(Optional Param As Boolean = False) As Boolean

    Static Result As Boolean
    Result = Param
    If (Param = False) Then Debug.Assert InIDE(True)
    InIDE = Result
    
End Function

Private Sub KillFile(sFile As String)
    If Len(Dir(sFile)) > 0 Then
        SetAttr sFile, vbNormal
        Kill sFile
    End If
End Sub
Private Function fnbMoveDirectoryFiles(ByRef sSourceDir As String, ByRef sDestDir As String) As Boolean
    On Error GoTo errTrap
    
    Dim sFile As String
    Dim lROCount As Long
    Dim lROIndex As Long
    Dim sROFiles() As String
    
    'check to see if source exists
    If Len(Dir(sSourceDir, vbDirectory)) = 0 Then
        MsgBox "Unable to find Directory :" & vbCr & vbCr & sSourceDir, vbCritical
        Exit Function
    End If

    'if destination doesn't exist, create it
    If Len(Dir(sDestDir, vbDirectory)) = 0 Then
        MkDir sDestDir
    End If

    lROCount = 0
    ReDim sROFiles(0 To lROCount)
    
    sFile = Dir(sSourceDir & "\*.*")
    Do While Len(sFile) > 0
        'Check if read only first
        If GetAttr(sSourceDir & "\" & sFile) And vbReadOnly Then
            lROCount = lROCount + 1
            ReDim Preserve sROFiles(0 To lROCount - 1)
          
            'Record which files are read only, so that we can make then read only afterwards...
            sROFiles(lROCount - 1) = sFile
        
            'Make it writable
            SetAttr sSourceDir & "\" & sFile, vbNormal
        End If
        
        'Rename all the files
        Name sSourceDir & "\" & sFile As sDestDir & "\" & sFile

    
        sFile = Dir
    Loop

    'Check to make sure that the directory is empty (to be sure , to be sure)
    If Len(Dir(sSourceDir & "\*.*")) > 0 Then
        MsgBox "Not all files have been moved from the Directory :" & vbCr & vbCr & sSourceDir, vbCritical
        Exit Function
    End If
    
    'Make those read only files
    For lROIndex = 0 To lROCount - 1
      SetAttr sDestDir & "\" & sROFiles(lROIndex), vbReadOnly
    Next lROIndex
    
    fnbMoveDirectoryFiles = True
    
    Exit Function
    
errTrap:
    MsgBox "fnbMoveDirectoryFiles Error: " & Err.Description & IIf(Erl, ", Line:" & Erl, "")
End Function

Private Sub ExitWithErrorLevel(ByVal lExitCode As Long)
    ' Call ExitProcess as the last action before closing
    ' otherwise it prevents proper clean up
    If Not InIDE Then
        ExitProcess lExitCode
    End If
End Sub

Private Function GetLongFilename(ByRef sShortName As String) As String
   
    Dim sArr()      As String
    sArr = Split(sShortName, "\")
   
    If UBound(sArr) >= 1 Then   '<volume>\<folder>\<file>
   
    Dim sPathSoFar      As String
    Dim sResult         As String
    Dim sShortPath      As String
        
    sPathSoFar = sArr(0)
    sShortPath = sPathSoFar
        
    Dim iCounter      As Integer
    Dim iMax          As Integer
    iMax = UBound(sArr)
        
    For iCounter = 1 To iMax
          
        sResult = Dir(sPathSoFar & "\", vbDirectory)
            
        Do
            
            If Len(sResult) = 0 Then
                'Path invalid
                GetLongFilename = vbNullString
                Exit Function
            End If
            
            If sResult = "." Or sResult = ".." Then
                sResult = Dir()
            Else
                    
                If UCase(GetShortFileName(sPathSoFar & "\" & sResult)) = UCase(sShortPath & "\" _
                    & sArr(iCounter)) Or UCase(sPathSoFar & "\" & sResult) = UCase(sShortPath & "\" _
                    & sArr(iCounter)) Then
                    Exit Do
                Else
                    sResult = Dir()
                End If
        
            End If
        
        Loop
        
        sShortPath = sShortPath & "\" & sArr(iCounter)
        sPathSoFar = sPathSoFar & "\" & sResult
        
    Next
        
    GetLongFilename = sPathSoFar
                
End If

End Function

 
Private Function GetShortFileName(ByRef LongPathName As String)

    Dim sShortPathName      As String
    Dim iBuffLen            As Integer
    Dim lRetVal             As Long
    
    sShortPathName = Space(255)
    iBuffLen = 255
    
    lRetVal = GetShortPathName(LongPathName, sShortPathName, iBuffLen)
    sShortPathName = Left(sShortPathName, lRetVal)

    GetShortFileName = sShortPathName
    
End Function


Private Function fnbCorrectPath32InProjectFile(ByRef sProjectFileName As String, ByRef sProjectDir As String, ByRef sOutputDir As String) As Boolean
    On Error GoTo errTrap

    Dim iInputFileNumber As Integer
    Dim iOutputFileNumber As Integer
    Dim sOriginalLine As String
    Dim sLine As String
    Dim sOutput As String
       
    'Open a new source file for writing
    iOutputFileNumber = FreeFile
        
    Open sProjectDir & "\NewPath32_" & Trim(sProjectFileName) For Output As iOutputFileNumber
    
    'Open the source file for reading
    iInputFileNumber = FreeFile
    Open sProjectDir & "\" & sProjectFileName For Input As iInputFileNumber
            
    Do Until EOF(iInputFileNumber)
        'Get a line from the file
        Line Input #iInputFileNumber, sOriginalLine
        
        'Trim any spaces from the beginning
        sLine = Trim(sOriginalLine)
        sOutput = sLine
        
        'Don't test an empty line
        If Len(sLine) > 0 Then
        
            If InStr(1, sLine, PATH32_LINE) = 1 Then
                If gbMaintainPaths Then
                    Dim sSplitPath() As String
                    sSplitPath = Split(gsOriginalOutDir, "\")
                    Dim sPrefix As String
                    Dim iCounter As Integer
                    For iCounter = 0 To UBound(sSplitPath) - 1
                        sPrefix = sPrefix & "..\"
                    Next
                    Dim iPathPos As Integer
                    iPathPos = InStr(sLine, "=")
                    sOutput = PATH32_LINE & Chr(34) & sPrefix & Mid(sLine, iPathPos + 2)
                End If
            End If
        
        End If
        
        'Output the line
        Print #iOutputFileNumber, sOutput
        
        
    Loop
    
    Close iInputFileNumber
    Close iOutputFileNumber
    
    On Error Resume Next
    Name sProjectDir & "\" & sProjectFileName As sProjectDir & "\OldPath32_" & sProjectFileName
    Name sProjectDir & "\NewPath32_" & Trim(sProjectFileName) As sProjectDir & "\" & sProjectFileName
    If Err.Number = 0 Then
        Kill sProjectDir & "\OldPath32_" & sProjectFileName
    Else
        Name sProjectDir & "\OldPath32_" & sProjectFileName As sProjectDir & "\" & sProjectFileName
        Kill sProjectDir & "\NewPath32_" & sProjectFileName
        MsgBox "Could not update PATH32 in line numbered project file", vbExclamation
    End If
    
    fnbCorrectPath32InProjectFile = True
    
    Exit Function
    
errTrap:
    MsgBox "fnbCorrectPath32InProjectFile Error: " & Err.Description & IIf(Erl, ", Line:" & Erl, "")
End Function

