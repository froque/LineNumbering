VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Line Numbering Tool Options"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Options.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Dirty"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3900
      TabIndex        =   26
      Top             =   60
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.Frame fraTabPages 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5355
      Index           =   0
      Left            =   300
      TabIndex        =   20
      Top             =   540
      Width           =   6375
      Begin VB.TextBox txtConditionalArgs 
         Height          =   345
         Left            =   240
         TabIndex        =   29
         Text            =   "Text2"
         ToolTipText     =   "Directory that line numbered source will be placed in"
         Top             =   4800
         Width           =   5775
      End
      Begin VB.CheckBox chkAutoIncrement 
         Alignment       =   1  'Right Justify
         Caption         =   "&Auto Increment"
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   4020
         Width           =   1515
      End
      Begin VB.TextBox txtRevision 
         Height          =   315
         Left            =   1740
         TabIndex        =   11
         ToolTipText     =   "Project will be lined numbered in these increments"
         Top             =   3660
         Width           =   615
      End
      Begin VB.TextBox txtMinor 
         Height          =   315
         Left            =   1020
         TabIndex        =   10
         ToolTipText     =   "Project will be lined numbered in these increments"
         Top             =   3660
         Width           =   615
      End
      Begin VB.TextBox txtMajor 
         Height          =   315
         Left            =   300
         TabIndex        =   9
         ToolTipText     =   "Project will be lined numbered in these increments"
         Top             =   3660
         Width           =   615
      End
      Begin VB.CheckBox chkVersionChanged 
         Caption         =   "Update &Version Number"
         Height          =   315
         Left            =   300
         TabIndex        =   8
         Top             =   3360
         Width           =   2475
      End
      Begin VB.CheckBox chkRetainPaths 
         Caption         =   "&Build in Same Place as Original"
         Height          =   315
         Left            =   3420
         TabIndex        =   15
         ToolTipText     =   "If set then the build path and compatible exe path will be modified in the new project to be the same as the original project"
         Top             =   3960
         Width           =   2535
      End
      Begin VB.TextBox txtCompileDirectory 
         Height          =   345
         Left            =   300
         TabIndex        =   5
         Text            =   "Text2"
         ToolTipText     =   "Directory that VB project will be built into"
         Top             =   2040
         Width           =   5775
      End
      Begin VB.TextBox txtLineNumbers 
         Height          =   315
         Left            =   300
         TabIndex        =   7
         Text            =   "Text3"
         ToolTipText     =   "Project will be lined numbered in these increments"
         Top             =   2880
         Width           =   1395
      End
      Begin VB.TextBox txtOutput 
         Height          =   345
         Left            =   300
         TabIndex        =   3
         Text            =   "Text2"
         ToolTipText     =   "Directory that line numbered source will be placed in"
         Top             =   1260
         Width           =   5775
      End
      Begin VB.TextBox txtProject 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Text1"
         ToolTipText     =   "Path and filename of VB project that will be operated on"
         Top             =   480
         Width           =   5775
      End
      Begin VB.CheckBox chkClearOutFile 
         Caption         =   "Clea&r Output Directory"
         Height          =   315
         Left            =   3420
         TabIndex        =   13
         ToolTipText     =   "De;ete all files in output directory before line numbering"
         Top             =   3060
         Width           =   2175
      End
      Begin VB.CheckBox chkBuildProject 
         Caption         =   "Build Pro&ject"
         Height          =   315
         Left            =   3420
         TabIndex        =   14
         ToolTipText     =   "Build VB project as well as line number"
         Top             =   3510
         Width           =   1635
      End
      Begin VB.Label label1 
         Caption         =   "Conditional Compilation Arguments"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   30
         Top             =   4560
         Width           =   4935
      End
      Begin VB.Label label1 
         Caption         =   "."
         Height          =   195
         Index           =   5
         Left            =   1680
         TabIndex        =   28
         Top             =   3780
         Width           =   195
      End
      Begin VB.Label label1 
         Caption         =   "."
         Height          =   195
         Index           =   4
         Left            =   960
         TabIndex        =   27
         Top             =   3780
         Width           =   195
      End
      Begin VB.Label label1 
         Caption         =   "B&uild Directory (Optional)"
         Height          =   255
         Index           =   3
         Left            =   300
         TabIndex        =   4
         Top             =   1800
         Width           =   4935
      End
      Begin VB.Label label1 
         Caption         =   "&Line Number Increments"
         Height          =   255
         Index           =   2
         Left            =   300
         TabIndex        =   6
         Top             =   2640
         Width           =   2715
      End
      Begin VB.Label label1 
         Caption         =   "Output &Directory (must be sub directory of project path)"
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   2
         Top             =   1020
         Width           =   4935
      End
      Begin VB.Label label1 
         Caption         =   "&Project Path"
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   0
         Top             =   240
         Width           =   2715
      End
   End
   Begin MSComctlLib.TabStrip tbsPageTabs 
      Height          =   5955
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   10504
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Options"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picFormBottom 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   6855
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6105
      Width           =   6915
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next >"
         Height          =   375
         Left            =   1200
         TabIndex        =   25
         Top             =   150
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "< &Previous"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   150
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame fraDivider 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   23
         Top             =   60
         Width           =   675
      End
      Begin VB.PictureBox picButtons 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   60
         ScaleHeight     =   375
         ScaleWidth      =   6855
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   60
         Width           =   6915
         Begin VB.CommandButton cmdApply 
            Caption         =   "&Save as Default"
            Height          =   375
            Left            =   3000
            TabIndex        =   16
            ToolTipText     =   "Saves these settings to the registry so they are the new defaults"
            Top             =   0
            Width           =   1635
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "&OK"
            Default         =   -1  'True
            Height          =   375
            Left            =   4740
            TabIndex        =   17
            ToolTipText     =   "Line number (and build if specified)"
            Top             =   0
            Width           =   975
         End
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
            Caption         =   "&Cancel"
            Height          =   375
            Left            =   5820
            TabIndex        =   18
            ToolTipText     =   "Cancel line numbering"
            Top             =   0
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mb_CancelPressed As Boolean  'True if user pressed cancel
Private mb_Dirty         As Boolean  'Used internally DIRTY state should be checked using Dirty property
Private mi_PreviouslyVisibleFrameIndex As Integer 'last selected tab page
Private mb_Loading      As Boolean
Private Const SPACING_TABFRAMES_LEFT = 150
Private Const SPACING_TABFRAMES_TOP = 400

Private Property Let Dirty(pb_Dirty As Boolean)
'Dirty Property - used to control forms dirty state

    mb_Dirty = pb_Dirty
    cmdApply.Enabled = mb_Dirty
    
End Property
Private Property Get Dirty() As Boolean
'Dirty Property - used to control forms dirty state

    Dirty = mb_Dirty
    
End Property

Private Sub Check1_Click()


    Dirty = (Check1.Value = vbChecked)

End Sub

Private Sub chkAutoIncrement_Click()

    Dirty = True
    
End Sub

Private Sub chkBuildProject_Click()

    Dirty = True

End Sub

Private Sub chkClearOutFile_Click()

    Dirty = True

End Sub

Private Sub chkRetainPaths_Click()

    Dirty = True

End Sub

Private Sub chkVersionChanged_Click()

    Dim bLocked As Boolean
    bLocked = Not (chkVersionChanged.Value = vbChecked)
    
    txtMajor.Locked = bLocked
    txtMinor.Locked = bLocked
    txtRevision.Locked = bLocked
    chkAutoIncrement.Enabled = Not bLocked
    If bLocked Then
        txtMajor.BackColor = vbButtonFace
        txtMinor.BackColor = vbButtonFace
        txtRevision.BackColor = vbButtonFace
    Else
        txtMajor.BackColor = vbWindowBackground
        txtMinor.BackColor = vbWindowBackground
        txtRevision.BackColor = vbWindowBackground
    End If
        
    Dirty = True
    
End Sub

Private Sub cmdApply_Click()
    
    Save True
    
End Sub

Private Sub cmdCancel_Click()
    
    mb_CancelPressed = True
    Unload Me
    
End Sub


Private Sub cmdOK_Click()

    mb_CancelPressed = False
    Unload Me
    
End Sub


Private Sub cmdNext_Click()

    tbsPageTabs.Tabs(tbsPageTabs.SelectedItem.Index + 1).Selected = True

End Sub

Private Sub cmdPrevious_Click()
    
    tbsPageTabs.Tabs(tbsPageTabs.SelectedItem.Index - 1).Selected = True
    
End Sub

Private Sub Form_Load()

    picButtons.BorderStyle = vbBSNone
    picFormBottom.BorderStyle = vbBSNone
    Dim lo_TabPage As Frame
    
    'Just in case - hide all tab frames
    For Each lo_TabPage In fraTabPages
        lo_TabPage.Visible = False
    Next

    'Set the first tab up...
    mi_PreviouslyVisibleFrameIndex = -1
    SetTabPage

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = vbFormControlMenu Then
        mb_CancelPressed = True
    End If
    
    Dim lb_Save As Boolean
    Dim lb_Permanently As Boolean
    
    'Do we need to save any changes?
    If Dirty Then
        
        If Not mb_CancelPressed Then
            lb_Save = True 'User pressed OK, so don't prompt to save
        Else
'            'User pressed X or Cancel so prompt to save changes?
'            Dim li_Response As Integer
'
'            li_Response = MsgBox("Save changes as new defaults?", vbYesNoCancel + vbDefaultButton1 + vbQuestion)
'
'            Select Case li_Response
'            Case vbYes
'                mb_CancelPressed = False
'                lb_Save = True
'                lb_Permanently = True
'            Case vbCancel
'                Cancel = True
'            End Select
        End If
        
    End If
    
    If lb_Save Then
        
        'Do the save - checking it worked
        If Not Save(lb_Permanently) Then
            'Save failed...so keep form open
            Cancel = True
        End If
        
    End If

End Sub

Private Sub Form_Resize()

    ResizeControls
    
End Sub

Private Sub picFormBottom_Resize()
    
    ResizeControls
    
End Sub

Private Sub ResizeControls()

    Const DISTANCE_FROM_TOP = 25
    Const DISTANCE_FROM_SIDES = 100
    
    picButtons.Left = picFormBottom.Width - picButtons.Width
    picButtons.Top = ((picFormBottom.Height - picButtons.Height) / 2) + DISTANCE_FROM_TOP
    
    'make just the bottom of the frame visible
    fraDivider.Width = picFormBottom.Width + 2 * DISTANCE_FROM_SIDES
    fraDivider.Left = -DISTANCE_FROM_SIDES
    fraDivider.Top = -fraDivider.Height + DISTANCE_FROM_TOP

    'Size Tabs
    tbsPageTabs.Width = Width - (tbsPageTabs.Left * 3)
    tbsPageTabs.Height = picFormBottom.Top - tbsPageTabs.Top - tbsPageTabs.Top
    
    Dim lo_TabPage As Frame
    For Each lo_TabPage In fraTabPages
        lo_TabPage.Width = tbsPageTabs.Width - (SPACING_TABFRAMES_LEFT * 2)
        lo_TabPage.Height = tbsPageTabs.Height - (SPACING_TABFRAMES_TOP * 1.5)
    Next

End Sub

Private Sub tbsPageTabs_BeforeClick(Cancel As Integer)
 
    'Store the current selection so it can be hidden later...
    mi_PreviouslyVisibleFrameIndex = tbsPageTabs.SelectedItem.Index - 1 'frame index is 0-2, tabs 1-3
        
End Sub

Private Sub tbsPageTabs_Click()

    SetTabPage

End Sub

Private Sub SetTabPage()

    Dim li_VisibleFrameIndex As Integer

    With tbsPageTabs
        li_VisibleFrameIndex = .SelectedItem.Index - 1 'frame index is 0-2, tabs 1-3
        
        With fraTabPages
            .Item(li_VisibleFrameIndex).Top = tbsPageTabs.Top + SPACING_TABFRAMES_TOP
            .Item(li_VisibleFrameIndex).Left = tbsPageTabs.Left + SPACING_TABFRAMES_LEFT
            .Item(li_VisibleFrameIndex).Visible = True
            
            On Error Resume Next    'for startup case where mi_PreviouslyVisibleFrameIndex = -1
            .Item(mi_PreviouslyVisibleFrameIndex).Visible = False
        End With
        
        cmdNext.Enabled = .Tabs.Count - 1 > li_VisibleFrameIndex
        cmdPrevious.Enabled = li_VisibleFrameIndex > 0
        
    End With

End Sub

Public Function ShowDialog(ParamArray pv_Arguments()) As Boolean
             
    If Not PopulateForm Then
        'If the form cannot be populated at the start then don't
        'bother to contine...
        ShowDialog = False
        Unload Me
    End If

    Show vbModal   'code stops here until form has been made invisible/closed
                    'activate event will fire next

    'TODO set any output parameters here
    
    'Close form .... This returns TRUE if okay pressed and FALSE if not
    ShowDialog = Not mb_CancelPressed
    Unload Me
        
End Function

Private Function Save(Optional ByVal Permanently As Boolean) As Boolean

    'TO DO: Write a save function
    'on Success set Dirty = False, Save = True on Fail set Save = False
    
    If Len(txtLineNumbers) = 0 Or Not IsNumeric(txtLineNumbers) _
        Or Val(txtLineNumbers) <= 0 Or CInt(Val(txtLineNumbers)) <> _
        CDbl(Val(txtLineNumbers)) Then
        
        MsgBox "Line number increment must whole number and greater than zero.", vbExclamation
        Exit Function
    
    End If
    
    gbMaintainPaths = (chkRetainPaths.Value = vbChecked)
    gbCompileProject = (chkBuildProject.Value = vbChecked)
    gbClearOutputDir = (chkClearOutFile.Value = vbChecked)
    gsProject = txtProject
    gsOutputDir = txtOutput
    gsCompileDir = txtCompileDirectory
    glIncrement = Val(txtLineNumbers)
    gbChangeVersion = (chkVersionChanged.Value = vbChecked)
    If gbChangeVersion Then
        gsMajor = Val(txtMajor)
        gsMinor = Val(txtMinor)
        gsRevision = Val(txtRevision)
        gbAutoIncrement = (chkAutoIncrement.Value = vbChecked)
    End If
    gsConditionalArgs = txtConditionalArgs
    
    If Permanently Then
        Dim sCommandString As String
        sCommandString = "C:\\Program Files\\Line Numbering\\linenumbering.exe /P%1 /I" _
            & txtLineNumbers & " /D " & IIf(gbChangeVersion, " /V" & _
            gsMajor & "." & gsMinor & "." & gsRevision & IIf(gbAutoIncrement, "y", "n"), "") & IIf(gbClearOutputDir, " /W ", "") & _
            " /O" & gsOutputDir & IIf(gbMaintainPaths, " /M ", "") & _
            IIf(gbCompileProject, " /C", "") & IIf(Len(gsCompileDir), gsCompileDir, "") & _
            " /T " & gsConditionalArgs
        Reg_SetRegistryValue HKEY_CLASSES_ROOT, "VisualBasic.Project\shell\Line Number VB Project\command", "", sCommandString, REG_SZ
    End If
    
    Dirty = False
    Save = True
    
End Function

Private Function PopulateForm() As Boolean

    'TO DO: Write populate form function
    'on Success set Dirty = False, PopulateForm = True on Fail set PopulateForm = False
    
    mb_Loading = True
    
    chkBuildProject.Value = IIf(gbCompileProject, vbChecked, vbUnchecked)
    chkClearOutFile.Value = IIf(gbClearOutputDir, vbChecked, vbUnchecked)
    chkRetainPaths.Value = IIf(gbMaintainPaths, vbChecked, vbUnchecked)
    txtProject = GetLongFilename(gsProject)
    txtOutput = gsOutputDir
    txtLineNumbers = glIncrement
    txtCompileDirectory = gsCompileDir
    chkVersionChanged = IIf(gbChangeVersion, vbChecked, vbUnchecked)
    chkVersionChanged_Click
    If gbChangeVersion Then
        txtMajor = gsMajor
        txtMinor = gsMinor
        txtRevision = gsRevision
        chkAutoIncrement = IIf(gbAutoIncrement, vbChecked, vbUnchecked)
    End If
    txtConditionalArgs = gsConditionalArgs
    
    mb_Loading = False
    
    Dirty = False
    PopulateForm = True
    
End Function



Private Sub txtCompileDirectory_Change()

    Dirty = True

End Sub

Private Sub txtConditionalArgs_Change()

    Dirty = True

End Sub

Private Sub txtLineNumbers_Change()

    Dirty = True

End Sub

Private Sub txtMajor_Change()

    Dirty = True
    
End Sub

Private Sub txtMinor_Change()

    Dirty = True
    
End Sub

Private Sub txtOutput_Change()

    Dirty = True

End Sub

Private Sub txtProject_Change()

    Dirty = True

End Sub

Private Sub txtRevision_Change()

    Dirty = True
    
End Sub
