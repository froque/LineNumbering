Attribute VB_Name = "Shell32"
Option Explicit

Private Declare Function OpenProcess Lib "kernel32" _
    (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = -1&

Public Sub ShellAndWaitToFinish( _
    ByVal EXEPath As String, _
    Optional ByVal WindowMode As VbAppWinStyle = vbMinimizedNoFocus, _
    Optional ByVal TimeOut As Long = INFINITE)
   
    Dim lProcessID      As Long
    Dim lProcess        As Long
    Dim lReturn         As Long
   
    On Error Resume Next
    lProcessID = Shell(EXEPath, CLng(WindowMode))
    
    If Err.Number <> 0 Then
        Exit Sub
    End If
    
    On Error GoTo 0

    lProcess = OpenProcess(SYNCHRONIZE, False, lProcessID)
    lReturn = WaitForSingleObject(lProcess, CLng(TimeOut))
    Call CloseHandle(lProcess)
   
End Sub
