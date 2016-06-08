Option Strict On
Option Explicit On
Imports System.IO
Imports VB = Microsoft.VisualBasic

Module MLineNumbering

    Private Const LINENUMBER_SIZE As Integer = 8
    Private Const LINE_CONTINUATION As String = "_"

    Private Const END_SUB As String = "End Sub"
    Private Const END_FUNCTION As String = "End Function"
    Private Const END_PROPERTY As String = "End Property"

    Private Const SUB_LINE As String = "Sub "
    Private Const PUBLIC_SUB As String = "Public Sub "
    Private Const Private_SUB As String = "Private Sub "
    Private Const FRIEND_SUB As String = "Friend Sub "

    Private Const FUNCTION_LINE As String = "Function "
    Private Const PUBLIC_FUNCTION As String = "Public Function "
    Private Const PRIVATE_FUNCTION As String = "Private Function "
    Private Const FRIEND_FUNCTION As String = "Friend Function "

    Private Const PROPERTY_LINE As String = "Property "
    Private Const PUBLIC_PROPERTY As String = "Public Property "
    Private Const PRIVATE_PROPERTY As String = "Private Property "
    Private Const FRIEND_PROPERTY As String = "Friend Property "

    Private Const MODULE_LINE As String = "Module="
    Private Const CLASS_LINE As String = "Class="
    Private Const USERCONTROL_LINE As String = "UserControl="
    Private Const DESIGNER_LINE As String = "Designer="
    Private Const FORM_LINE As String = "Form="
    Private Const RELATEDDOC_LINE As String = "RelatedDoc="
    Private Const RESFILE_LINE As String = "ResFile32="
    Private Const COMPATIBLEEXE32_LINE As String = "CompatibleEXE32="
    Private Const PATH32_LINE As String = "Path32="
    Private Const MAJORVER_LINE As String = "MajorVer="
    Private Const MINORVER_LINE As String = "MinorVer="
    Private Const REVISIONVER_LINE As String = "RevisionVer="
    Private Const AUTOINCREMENTVER_LINE As String = "AutoIncrementVer="

    Private Const CASE_STATEMENT As String = "Case"
    Private Const SELECTCASE_STATEMENT As String = "Select Case"
    Private Const ENDSELECT_STATEMENT As String = "End Select"
    Private Const CONDITIONAL_STATEMENT As String = "#"
    Private Const COMMENT_STATEMENT As String = "'"
    Private Const REM_STATEMENT As String = "Rem"

    Public Sub Main(args As String())
        Try
            Dim options As New Options()
            If Not CommandLine.Parser.Default.ParseArguments(args, options) Then
                Environment.Exit(1)
            End If

            Dim config As New Config
            config.Project = New FileInfo(options.Project)
            config.Output = New DirectoryInfo(Path.Combine(config.Project.Directory.FullName, options.Output))
            config.Increment = options.Increment

            'create output directory
            config.Output.Create()
            deleteFilesFromFolder(config.Output)

            parseProjectFile(config)

            Console.WriteLine("Completed Successfully.")
            Environment.Exit(0)
            Exit Sub
        Catch ex As Exception
            Console.WriteLine(ex.ToString())
            Environment.Exit(-1)
        End Try

    End Sub

    Sub deleteFilesFromFolder(folder As DirectoryInfo)
        If folder.Exists Then
            For Each _file As FileInfo In folder.GetFiles()
                _file.Delete()
            Next
            For Each _folder As DirectoryInfo In folder.GetDirectories()
                deleteFilesFromFolder(_folder)
            Next
        End If
    End Sub

    Private Sub parseProjectFile(ByVal config As Config)
        Dim iInputFileNumber As Integer
        Dim iOutputFileNumber As Integer


        'Open a new source file for writing
        iOutputFileNumber = FreeFile()
        FileOpen(iOutputFileNumber, Path.Combine(config.Output.FullName, config.Project.Name), OpenMode.Output)

        'Open the source file for reading
        iInputFileNumber = FreeFile()
        FileOpen(iInputFileNumber, config.Project.FullName, OpenMode.Input)

        Do Until EOF(iInputFileNumber)
            Dim sOriginalLine As String
            Dim sOutput As String

            'Get a line from the file
            sOriginalLine = LineInput(iInputFileNumber)

            sOutput = processProjectLine(config, sOriginalLine)

            'Output the line
            PrintLine(iOutputFileNumber, sOutput)

        Loop

        FileClose(iInputFileNumber)
        FileClose(iOutputFileNumber)
    End Sub

    Private Function processProjectLine(ByVal config As Config, ByVal sOriginalLine As String) As String
        Dim iFilePos As Integer
        Dim sLine As String
        Dim sOutput As String
        Dim bGetFile As Boolean
        Dim bCopyFile As Boolean
        Dim bParseFile As Boolean
        Dim bCopyRenameFile As Boolean
        Dim sFile As String
        Dim sFileName As String
        Dim bCheckFRX As Boolean
        Dim bCheckCTX As Boolean
        Dim bCheckDesigner As Boolean
        Dim bAutoInc As Boolean
        Dim sTemp As String
        Dim sProjectDir As String
        Dim sOutputDir As String

        sProjectDir = config.Project.DirectoryName
        sOutputDir = config.Output.FullName

        bGetFile = False 'We need to process the file
        bParseFile = False 'We need to add line numbers
        bCopyFile = False 'Copy File only
        bCopyRenameFile = False 'Copy and rename (used for compatible)
        bCheckFRX = False
        bCheckCTX = False
        bCheckDesigner = False

        'Trim any spaces from the beginning
        sLine = Trim(sOriginalLine)
        sOutput = sLine

        'Don't test an empty line
        If Len(sLine) = 0 Then
            processProjectLine = ""
            Exit Function
        End If

        'Check for Forms,Modules,UserControls or Classes
        If InStr(1, sLine, MODULE_LINE) = 1 Or InStr(1, sLine, CLASS_LINE) = 1 Then
            bGetFile = True
            bParseFile = True
        ElseIf InStr(1, sLine, USERCONTROL_LINE) = 1 Then
            bGetFile = True
            bParseFile = True
            bCheckCTX = True
        ElseIf InStr(1, sLine, DESIGNER_LINE) = 1 Then
            bGetFile = True
            bParseFile = True
            bCheckDesigner = True
        ElseIf InStr(1, sLine, FORM_LINE) = 1 Then
            bGetFile = True
            bParseFile = True
            bCheckFRX = True
            'Now check for Related Documents and Resource File
        ElseIf InStr(1, sLine, RELATEDDOC_LINE) = 1 Or InStr(1, sLine, RESFILE_LINE) = 1 Then
            bGetFile = True
            bCopyFile = True
            'Now adjust the CompatibleExe if required...
        ElseIf InStr(1, sLine, COMPATIBLEEXE32_LINE) = 1 Then
            bGetFile = True
            bCopyRenameFile = True
        ElseIf InStr(1, sLine, PATH32_LINE) = 1 Then
            sOutput = PATH32_LINE & Chr(34) & Chr(34)
            If config.maintainPaths Then
                sOutput = sLine
            End If
        ElseIf InStr(1, sLine, MAJORVER_LINE) = 1 Then
            iFilePos = InStr(1, sLine, "=")
            If config.changeVersion Then
                sOriginalLine = MAJORVER_LINE & config.major
                sOutput = MAJORVER_LINE & config.major
            End If
        ElseIf InStr(1, sLine, MINORVER_LINE) = 1 Then
            iFilePos = InStr(1, sLine, "=")
            If config.changeVersion Then
                sOriginalLine = MINORVER_LINE & config.minor
                sOutput = MINORVER_LINE & config.minor
            End If
        ElseIf InStr(1, sLine, REVISIONVER_LINE) = 1 Then
            iFilePos = InStr(1, sLine, "=")
            If config.changeVersion Then
                If config.autoIncrement Then
                    sOriginalLine = REVISIONVER_LINE & (Val(config.revision) + 1)
                Else
                    sOriginalLine = REVISIONVER_LINE & config.revision
                End If
                sOutput = REVISIONVER_LINE & config.revision
            Else
                sTemp = Trim(Mid(sLine, iFilePos + 1))
                If IsNumeric(sTemp) Then
                    sOriginalLine = REVISIONVER_LINE & CStr(Val(sTemp) + 1)
                End If
            End If
        ElseIf InStr(1, sLine, AUTOINCREMENTVER_LINE) = 1 Then
            If config.changeVersion Then
                sOutput = AUTOINCREMENTVER_LINE & (IIf(config.autoIncrement, "1", "0")).ToString
                sOriginalLine = AUTOINCREMENTVER_LINE & (IIf(config.autoIncrement, "1", "0")).ToString
            Else
                iFilePos = InStr(1, sLine, "=")
                sTemp = Trim(Mid(sLine, iFilePos + 1))
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
                If Left(sFile, 1) = """" And Right(sFile, 1) = """" Then
                    sFile = Mid(sFile, 2, Len(sFile) - 2)
                End If


                'Get just the File Name
                If InStr(sFile, "\") > 0 Then
                    sFileName = Right(sFile, Len(sFile) - InStrRev(sFile, "\"))
                    sFile = sProjectDir & "\" & sFile
                Else
                    sFileName = sFile
                    sFile = sProjectDir & "\" & sFileName
                End If

                sOutput = Left(sLine, iFilePos - 1) & sFileName
                If Len(sFile) <> 0 Then
                    If bParseFile Then
                        'For code add the numbers
                        If AddLineNumbers(config, sFile, sOutputDir) = False Then
                            Throw New Exception("Unable to add line numbers to File :" + sFile)
                        End If
                    ElseIf bCopyFile Then
                        'Just copt the file over
                        If Dir(sFile) = "" Then
                            Throw New Exception("Unable to add line numbers to File :" + sFile)
                        End If
                        FileCopy(sFile, sOutputDir & "\" & sFileName)
                    ElseIf bCopyRenameFile Then
                        'Copy File over and rename it (to avoid conflicts)
                        If Left(sFile, 2) = ".." Then
                            sFile = sProjectDir & "\" & sFile
                        End If

                        If Dir(sFile) = "" Then
                            Throw New Exception("Unable to add line numbers to File :" + sFile)
                        End If

                        sOutput = Left(sLine, iFilePos - 1) & "CMP_" & sFileName
                        FileCopy(sFile, sOutputDir & "\" & "CMP_" & sFileName)
                    End If

                    If bCheckFRX Then
                        'If the file is a form, check for an FRX
                        If UCase(Right(sFile, 3)) = "FRM" Then
                            sFile = Left(sFile, Len(sFile) - 3) & "frx"
                            sFileName = Left(sFileName, Len(sFileName) - 3) & "frx"

                            If Len(Dir(sFile)) > 0 Then
                                FileCopy(sFile, sOutputDir & "\" & sFileName)
                            End If
                        End If
                    End If

                    If bCheckDesigner Then
                        'If the file is a form, check for an FRX
                        If UCase(Right(sFile, 3)) = "DSR" Then
                            sFile = Left(sFile, Len(sFile) - 3) & "DCA"
                            sFileName = Left(sFileName, Len(sFileName) - 3) & "DCA"

                            If Len(Dir(sFile)) > 0 Then
                                FileCopy(sFile, sOutputDir & "\" & sFileName)
                            End If

                            sFile = Left(sFile, Len(sFile) - 3) & "dsx"
                            sFileName = Left(sFileName, Len(sFileName) - 3) & "Dsx"

                            If Len(Dir(sFile)) > 0 Then
                                FileCopy(sFile, sOutputDir & "\" & sFileName)
                            End If
                        End If
                    End If

                    If bCheckCTX Then
                        'If the file is a user control, check for an CTX
                        If UCase(Right(sFile, 3)) = "CTL" Then
                            sFile = Left(sFile, Len(sFile) - 3) & "ctx"
                            sFileName = Left(sFileName, Len(sFileName) - 3) & "ctx"

                            If Len(Dir(sFile)) > 0 Then
                                FileCopy(sFile, sOutputDir & "\" & sFileName)
                            End If
                        End If
                    End If
                End If

            End If
        End If

        processProjectLine = sOutput
    End Function


    Private Function AddLineNumbers(ByVal config As Config, ByRef sFile As String, ByRef sOutputDir As String) As Boolean

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
        Dim sLineNumberStr As String
        Dim sFirstToken As String
        Dim bFoundNumbers As Boolean

        'Retrieve the File Name
        Dim sOldFileName As String
        sOldFileName = sFile

        Dim sFileInfo As FileInfo = New FileInfo(sFile)
        sFileName = sFileInfo.Name
        sFileDir = sFileInfo.Directory.FullName

        Console.WriteLine("adding line numbers to file: " & sFileName)

        'Open a new source file for writing
        iOutputFileNumber = FreeFile()
        FileOpen(iOutputFileNumber, sOutputDir & "\" & sFileName, OpenMode.Output)

        'Open the source file for reading
        iInputFileNumber = FreeFile()
        FileOpen(iInputFileNumber, sFileDir & "\" & sFileName, OpenMode.Input)

        'Reset the flags
        bInProc = False
        iLineNumberCount = config.Increment

        'Loop through the file
        Dim lNumberChars As Integer
        Dim lExtraChars As Integer
        Dim lPosToken As Integer
        Do While Not EOF(iInputFileNumber)

            sLine = LineInput(iInputFileNumber)

            ' convert tabs to spaces first
            sTrimmedLine = Replace(sLine, vbTab, " ")
            sTrimmedLine = Trim(sTrimmedLine)

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
                If Right(sTrimmedLine, 1) = LINE_CONTINUATION Then
                    bSkipNextLine = True
                End If

                'Are we leaving a procedure?
                If bInProc Then
                    If InStr(1, sTrimmedLine, END_SUB) = 1 Or InStr(1, sTrimmedLine, END_FUNCTION) = 1 Or InStr(1, sTrimmedLine, END_PROPERTY) = 1 Then
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
                            sLine = Space(LINENUMBER_SIZE) & sLine
                        End If
                    End If

                    If InStr(1, sTrimmedLine, CONDITIONAL_STATEMENT) = 1 Then
                        bSkipThisLine = True

                        'Add in spaces to keep justified...
                        If bFoundNumbers = False Then
                            sLine = Space(LINENUMBER_SIZE) & sLine
                        End If
                    End If

                    If InStr(1, sTrimmedLine, REM_STATEMENT) = 1 Or InStr(1, sTrimmedLine, COMMENT_STATEMENT) = 1 Then

                        'Not allowed to put line numbers between a select and the first case statement
                        If bStartOfSelect = True Then
                            bSkipThisLine = True

                            'Add in spaces to keep justified...
                            If bFoundNumbers = False Then
                                sLine = Space(LINENUMBER_SIZE) & sLine
                            End If
                        End If
                    End If

                    'Don't worry about these types of lines, compiles OK, but we keep the code justified...
                    'DIM_STATEMENT, STATIC_STATEMENT


                    'Are we entering a procedure?
                Else
                    If InStr(1, sTrimmedLine, SUB_LINE) = 1 Or InStr(1, sTrimmedLine, PUBLIC_SUB) = 1 Or InStr(1, sTrimmedLine, Private_SUB) = 1 Or InStr(1, sTrimmedLine, FRIEND_SUB) = 1 Or InStr(1, sTrimmedLine, FUNCTION_LINE) = 1 Or InStr(1, sTrimmedLine, PUBLIC_FUNCTION) = 1 Or InStr(1, sTrimmedLine, PRIVATE_FUNCTION) = 1 Or InStr(1, sTrimmedLine, FRIEND_FUNCTION) = 1 Or InStr(1, sTrimmedLine, PROPERTY_LINE) = 1 Or InStr(1, sTrimmedLine, PUBLIC_PROPERTY) = 1 Or InStr(1, sTrimmedLine, PRIVATE_PROPERTY) = 1 Or InStr(1, sTrimmedLine, FRIEND_PROPERTY) = 1 Then
                        'Inside Proc
                        bInProc = True
                        bSkipThisLine = True
                    End If
                End If

            End If

            'Are we in a procedure and not skipping this line
            If bInProc = True And bSkipThisLine = False Then
                'Add a line number
                sLineNumberStr = String.Format("{0,-8}", iLineNumberCount)

                'Check if any existing numbers
                sFirstToken = Split(sTrimmedLine, " ")(0)

                If IsNumeric(sFirstToken) Then
                    bFoundNumbers = True

                    sLine = sTrimmedLine


                    lNumberChars = Len(sFirstToken)
                    lExtraChars = LINENUMBER_SIZE - lNumberChars
                    lPosToken = InStr(sLine, sFirstToken)

                    'Trimming the Number off the line, and also trying to keep everything justified...
                    If lNumberChars <= LINENUMBER_SIZE Then
                        sLine = Mid(sLine, lPosToken + lNumberChars)
                        If Left(sLine, lExtraChars) = Space(lExtraChars) Then
                            sLine = Mid(sLine, lExtraChars + 1)
                        End If
                    Else
                        sLine = Mid(sLine, lPosToken + lNumberChars)
                    End If
                End If

                sLine = sLineNumberStr & sLine
                PrintLine(iOutputFileNumber, sLine)
                iLineNumberCount = iLineNumberCount + config.Increment
            Else
                PrintLine(iOutputFileNumber, sLine)
            End If
        Loop

        FileClose(iInputFileNumber)
        FileClose(iOutputFileNumber)

        AddLineNumbers = True

        Exit Function

    End Function


End Module