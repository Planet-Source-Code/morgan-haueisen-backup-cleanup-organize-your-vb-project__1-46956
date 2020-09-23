VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "VB Project BackUp"
   ClientHeight    =   6720
   ClientLeft      =   3315
   ClientTop       =   2280
   ClientWidth     =   6045
   ForeColor       =   &H00000000&
   Icon            =   "frmMainAddIn.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6720
   ScaleWidth      =   6045
   Begin VB.PictureBox picOptions 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000014&
      ForeColor       =   &H80000008&
      Height          =   3210
      Left            =   2025
      ScaleHeight     =   3180
      ScaleWidth      =   3450
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   975
      Visible         =   0   'False
      Width           =   3480
      Begin VB.CheckBox chkDeleteEmpty 
         BackColor       =   &H80000014&
         Caption         =   "Delete empty folders"
         Height          =   210
         Left            =   120
         TabIndex        =   24
         Top             =   1020
         Value           =   1  'Checked
         Width           =   3225
      End
      Begin VB.OptionButton optMove 
         BackColor       =   &H80000014&
         Caption         =   "Move"
         Height          =   225
         Left            =   135
         TabIndex        =   22
         Top             =   1725
         Width           =   1155
      End
      Begin VB.OptionButton optCopy 
         BackColor       =   &H80000014&
         Caption         =   "Copy"
         Height          =   225
         Left            =   135
         TabIndex        =   21
         Top             =   1455
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.CheckBox chkOrganize 
         BackColor       =   &H80000014&
         Caption         =   "Organize files into separate folders"
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Value           =   1  'Checked
         Width           =   3225
      End
      Begin VB.CheckBox chkOnlyNewer 
         BackColor       =   &H80000014&
         Caption         =   "Copy only newer files"
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   2430
         Width           =   3225
      End
      Begin VB.CheckBox chkPromptBefore 
         BackColor       =   &H80000014&
         Caption         =   "Prompt before overwriting existing files"
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   2130
         Width           =   3225
      End
      Begin VbProjectBackUp.chameleonButton cmdCLose 
         Height          =   285
         Left            =   75
         TabIndex        =   17
         Top             =   2820
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "Close Menu"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMainAddIn.frx":0CCA
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CheckBox chkSupportFile 
         BackColor       =   &H80000014&
         Caption         =   "Include support files in root folder"
         Height          =   210
         Left            =   120
         TabIndex        =   16
         Top             =   420
         Value           =   1  'Checked
         Width           =   3225
      End
      Begin VB.CheckBox chkSupportDir 
         BackColor       =   &H80000014&
         Caption         =   "Include support files in sub-folders"
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   135
         Value           =   1  'Checked
         Width           =   3225
      End
      Begin VB.Label lblCopyMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Files that are in the parent folder"
         Height          =   405
         Left            =   1440
         TabIndex        =   23
         Top             =   1485
         Width           =   1680
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         BorderStyle     =   5  'Dash-Dot-Dot
         Index           =   2
         X1              =   0
         X2              =   5000
         Y1              =   2010
         Y2              =   2010
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         Index           =   1
         X1              =   0
         X2              =   5000
         Y1              =   2745
         Y2              =   2745
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   5000
         Y1              =   1350
         Y2              =   1350
      End
   End
   Begin VB.PictureBox picToolBar 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   0
      ScaleHeight     =   1635
      ScaleWidth      =   6045
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   6045
      Begin VbProjectBackUp.chameleonButton cmdStart 
         Height          =   885
         Left            =   2919
         TabIndex        =   8
         ToolTipText     =   "Begin BackUp"
         Top             =   30
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Start Copy"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmMainAddIn.frx":0CE6
         PICN            =   "frmMainAddIn.frx":0D02
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VbProjectBackUp.chameleonButton cmdAbort 
         Height          =   885
         Left            =   2919
         TabIndex        =   9
         Top             =   30
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   1561
         BTYPE           =   3
         TX              =   "Abort"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmMainAddIn.frx":19DC
         PICN            =   "frmMainAddIn.frx":19F8
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtPath 
         Height          =   300
         Left            =   105
         TabIndex        =   6
         Top             =   1215
         Width           =   5370
      End
      Begin VbProjectBackUp.chameleonButton cmdAbout 
         Height          =   885
         Left            =   4182
         TabIndex        =   10
         Top             =   30
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "About"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmMainAddIn.frx":26D2
         PICN            =   "frmMainAddIn.frx":26EE
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VbProjectBackUp.chameleonButton cmdQuit 
         Height          =   885
         Left            =   5130
         TabIndex        =   11
         Top             =   30
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Exit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmMainAddIn.frx":2BD2
         PICN            =   "frmMainAddIn.frx":2BEE
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VbProjectBackUp.chameleonButton cmdOptions 
         Height          =   885
         Left            =   1971
         TabIndex        =   13
         Top             =   30
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   1561
         BTYPE           =   9
         TX              =   "Options"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmMainAddIn.frx":3213
         PICN            =   "frmMainAddIn.frx":322F
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VbProjectBackUp.chameleonButton cmdFindFolder 
         Height          =   330
         Left            =   5520
         TabIndex        =   12
         ToolTipText     =   "Browser"
         Top             =   1200
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "..."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   65280
         MPTR            =   1
         MICON           =   "frmMainAddIn.frx":36BE
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   1
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblMisc 
         BackStyle       =   0  'Transparent
         Caption         =   "Destination Folder:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1005
         Width           =   2505
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   6000
         Y1              =   975
         Y2              =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   6000
         Y1              =   960
         Y2              =   960
      End
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   0
      ScaleHeight     =   1140
      ScaleWidth      =   6045
      TabIndex        =   2
      Top             =   5580
      Width           =   6045
      Begin VB.PictureBox picProg 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10
         ScaleHeight     =   240
         ScaleWidth      =   5340
         TabIndex        =   4
         Top             =   -15
         Visible         =   0   'False
         Width           =   5400
      End
      Begin VB.Label lblPrinting 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   10
         TabIndex        =   3
         Top             =   360
         Width           =   5325
         WordWrap        =   -1  'True
      End
   End
   Begin VB.ListBox ProjectList 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      Height          =   840
      IntegralHeight  =   0   'False
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Project File Information"
      Top             =   2025
      Width           =   5700
   End
   Begin VB.Label lblProject 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFE1D5&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   275
      Left            =   0
      TabIndex        =   1
      Top             =   1695
      UseMnemonic     =   0   'False
      Width           =   6045
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**************************************/
'/*     Author: Morgan Haueisen        */
'/*             morganh@hartcom.net    */
'/*     Copyright (c) 2003-2004        */
'/**************************************/
'Legal:
'        This is intended for and was uploaded to www.planetsourcecode.com
'
'        Redistribution of this code, whole or in part, as source code or in binary form, alone or
'        as part of a larger distribution or product, is forbidden for any commercial or for-profit
'        use without the author's explicit written permission.
'
'        Redistribution of this code, as source code or in binary form, with or without
'        modification, is permitted provided that the following conditions are met:
'
'        Redistributions of source code must include this list of conditions, and the following
'        acknowledgment:
'
'        This code was developed by Morgan Haueisen.  <morganh@hartcom.net>
'        Source code, written in Visual Basic, is freely available for non-commercial,
'        non-profit use at www.planetsourcecode.com.
'
'        Redistributions in binary form, as part of a larger project, must include the above
'        acknowledgment in the end-user documentation.  Alternatively, the above acknowledgment
'        may appear in the software itself, if and wherever such third-party acknowledgments
'        normally appear.

Option Explicit

Private Const C_WarnMessage As String = "The source and destination paths are the same." & vbCrLf & vbCrLf & _
    "If you continue you will be modifying your original project(s) files; " & _
    "any old or obsolete files will not be moved or delete. If this is the first " & _
    "time running this utility, you may want to create a new folder to take advantage " & _
    "of the cleaning functionality it provides." & vbCrLf & vbCrLf
    
Private Const m_cWarnMessage2 As String = "You have selected the option to move files; " & _
    "at the completion of the process you need to close and re-open this project!" & vbCrLf & vbCrLf
    
    

Private mcProgBar                 As clsProgressBar
Private mcScreen                  As clsScreenSize
Private mcFile                    As clsFileUtilities
Private mcProgList()              As clsVirtualListbox

Private mstrPathNames(0 To 1, 1 To 11) As String
Private mstrProjectGroups()            As String

Private mblnQuitCommand           As Boolean
Private mstrActiveMessage         As String
Private mstrProjectFileName       As String
Private mstrProjectFilePath       As String
Private mstrNewProjectPath        As String
Private mblnReLoadProject         As Boolean
Private mblnIsVBG                 As Boolean
Private mblnOnlyNewer             As Boolean

Private Type ProjectInfoType
    PName                        As String
    MajorVer                     As String
    MinorVer                     As String
    RevisionVer                  As String
End Type
Private mudtProjectInfo           As ProjectInfoType

'/* Used to play system sounds
Private Const MC_iAsterisk       As Long = &H10&
Private Const MC_iQuestion       As Long = &H20&
Private Const MC_iExclamation    As Long = &H30&
Private Const MC_iInformation    As Long = &H40&
Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long

'/* Used for Manifest files (Win XP style controls)
Private Declare Sub InitCommonControls Lib "Comctl32.dll" ()


Private Sub chkOnlyNewer_Click()

    If chkOnlyNewer.Value = vbChecked Then
        chkPromptBefore.Value = vbUnchecked
    End If

End Sub

Private Sub chkPromptBefore_Click()

    If chkPromptBefore.Value = vbChecked Then
        chkOnlyNewer.Value = vbUnchecked
    End If

End Sub

Private Sub cmdAbort_Click()

    mblnQuitCommand = True

End Sub


Private Sub cmdStart_Click()

  Dim lngI As Long

    On Error GoTo Err_Proc

    txtPath.Text = Trim$(txtPath.Text)
    If txtPath.Text = vbNullString Then
        frmMsgBox.SMessageModal "No Destination Folder Selected", vbCritical, , , , False
        Exit Sub
    End If
    
    If LenB(Dir$(mstrProjectFileName)) = 0 Or LenB(mstrProjectFileName) = 0 Then
        frmMsgBox.SMessageModal "No Project File Found", vbCritical, , , , False
        Exit Sub
    End If

    cmdStart.Visible = False
    cmdOptions.Visible = False
    cmdQuit.Visible = False
    cmdAbout.Visible = False
    ProjectList.Enabled = False
    cmdAbort.Visible = True
    cmdAbort.Enabled = True
    picOptions.Visible = False
    mblnQuitCommand = False
    mblnReLoadProject = False
    
    If chkOrganize.Value = vbChecked Then
        mstrPathNames(1, 1) = "PropertyPage\"
        mstrPathNames(1, 2) = "UserControl\"
        mstrPathNames(1, 3) = "RelatedDoc\"
        mstrPathNames(1, 4) = "Designer\"
        mstrPathNames(1, 5) = "Class\"
        mstrPathNames(1, 6) = "Form\"
        mstrPathNames(1, 7) = "Module\"
        mstrPathNames(1, 8) = "RelatedDoc\"
        mstrPathNames(1, 9) = "UserDocument\"
        mstrPathNames(1, 10) = "Library\"
        mstrPathNames(1, 11) = "Documentation\"
    Else
        For lngI = 1 To UBound(mstrPathNames)
            mstrPathNames(1, lngI) = vbNullString
        Next lngI
    End If

    mstrNewProjectPath = txtPath.Text
    If right$(mstrNewProjectPath, 1) <> "\" Then mstrNewProjectPath = mstrNewProjectPath & "\"

    mcFile.NoConfirmation = Not CBool(chkPromptBefore.Value)
    mblnOnlyNewer = CBool(chkOnlyNewer.Value)

    If mblnIsVBG Then
        If Not CreateNewGroupProjectFile Then GoTo Exit_Here
    End If

    For lngI = 0 To UBound(mstrProjectGroups)
        If Not CreateNewProjectFile(lngI) Then Exit For
        Call CopyProjectFiles(lngI)
        If mblnQuitCommand Then Exit For
    Next lngI
    
    If Not mblnQuitCommand Then
        If Not mblnReLoadProject Then
            Call MessageBeep(MC_iInformation)
        Else
            frmMsgBox.SMessageModal m_cWarnMessage2 & "Do NOT save anything when exiting!", vbExclamation
        End If
    End If
    
Exit_Here:
    cmdAbort.Visible = False
    cmdStart.Visible = True
    cmdOptions.Visible = True
    cmdQuit.Visible = True
    cmdAbout.Visible = True
    ProjectList.Enabled = True
    picProg.Visible = False
    mblnQuitCommand = False

Exit Sub


Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "frmMain", "cmdStart_Click"
    Err.Clear
    Resume Exit_Here

End Sub

Private Function CreateNewProjectFile(ByVal vlngIndex As Long) As Boolean

  Dim lngFN               As Long
  Dim lngNFN              As Long
  Dim lngX                As Long
  Dim strLine             As String
  Dim strTemp             As String
  Dim strProjectName      As String
  Dim strProjectPath      As String
  Dim strNewProjectName   As String
  Dim strNewProjectPath   As String
  Dim strPathPrefix       As String
  Dim blnRefresh          As Boolean


    On Error GoTo Err_Proc

    If LenB(mstrProjectGroups(vlngIndex)) Then
        strTemp = mstrProjectGroups(vlngIndex)
     Else
        strTemp = mstrProjectFileName
    End If
    If strTemp = vbNullString Then Exit Function

    strProjectPath = mcFile.RetOnlyPath(strTemp)
    strProjectName = mcFile.RetOnlyFilename(strTemp)
    strNewProjectName = strProjectName
    strNewProjectPath = vbNullString
    blnRefresh = False

    strNewProjectPath = mstrNewProjectPath
    '/* Add Group path
    If LenB(mstrProjectGroups(vlngIndex)) Then
        strPathPrefix = mcFile.RetOnlyFilename(mstrProjectGroups(vlngIndex))
        strPathPrefix = left$(strPathPrefix, Len(strPathPrefix) - 4) & "\"
        strNewProjectPath = strNewProjectPath & strPathPrefix
    End If

    mcFile.CreateDir strNewProjectPath
    
    '/* Check for same folder
    If mcFile.RetShortPathName(strProjectPath) = mcFile.RetShortPathName(strNewProjectPath) Then
        '/* Make new file name
        strTemp = strNewProjectName
        strTemp = Mid$(strNewProjectName, 1, Len(strNewProjectName) - 4)
        If Len(strNewProjectName) > 14 Then
            If Val(right$(strNewProjectName, 14)) > 0 Then
                strTemp = left$(strTemp, Len(strTemp) - 14)
                strNewProjectName = strTemp & Format$(Now, "yyyymmddhhmmss") & Mid$(strNewProjectName, Len(strNewProjectName) - 3)
            Else
                strNewProjectName = strTemp & Format$(Now, "yyyymmddhhmmss") & Mid$(strNewProjectName, Len(strNewProjectName) - 3)
            End If
        Else
            strNewProjectName = strTemp & Format$(Now, "yyyymmddhhmmss") & Mid$(strNewProjectName, Len(strNewProjectName) - 3)
        End If
        
        '/* Ignore if a group project is loaded
        If Not mblnIsVBG Then
            If optMove.Value Then mstrActiveMessage = m_cWarnMessage2
            If frmMsgBox.SMessageModal(C_WarnMessage & mstrActiveMessage & "Do you wish to continue?", vbQuestion + vbYesNo, _
                "WARNING", , , , Me.Width, , Me) = vbNo Then
                
                CreateNewProjectFile = False
                GoTo Exit_Here
            End If
        End If
        
        blnRefresh = True
        cmdAbort.Enabled = False
        mblnReLoadProject = True
    
    End If

    '/* Begin building project file
    lngFN = FreeFile
    Open strProjectPath & strProjectName For Input As #lngFN

    lngNFN = FreeFile
    Open strNewProjectPath & strNewProjectName For Output As #lngNFN

    Do
        Line Input #lngFN, strLine
        If EOF(lngFN) Then Exit Do

        Select Case left$(strLine, 5)
         Case "Prope" 'PropertyPage=
            strTemp = Mid$(strLine, 14)
            strTemp = mcFile.RetOnlyFilename(strTemp)
            strLine = "PropertyPage=" & mstrPathNames(1, 1) & strTemp
         Case "UserC" 'UserControl=
            strTemp = Mid$(strLine, 13)
            strTemp = mcFile.RetOnlyFilename(strTemp)
            strLine = "UserControl=" & mstrPathNames(1, 2) & strTemp
         Case "UserD" 'UserDocument=
            strTemp = Mid$(strLine, 14)
            strTemp = mcFile.RetOnlyFilename(strTemp)
            strLine = "UserDocument=" & mstrPathNames(1, 9) & strTemp
         Case "ResFi" 'ResFile32="XXXX.res"
            strTemp = Mid$(strLine, 11)
            strTemp = Mid$(strTemp, 2, Len(strTemp) - 2)
            strTemp = mcFile.RetOnlyFilename(strTemp)
            strLine = "ResFile32=" & Chr$(34) & mstrPathNames(1, 3) & strTemp & Chr$(34)
         Case "Desig" 'Designer=
            strTemp = Mid$(strLine, 10)
            strTemp = mcFile.RetOnlyFilename(strTemp)
            strLine = "Designer=" & mstrPathNames(1, 4) & strTemp
         Case "Class" 'Class=xxxxx; xxxxx.cls
            lngX = InStr(strLine, " ")
            If lngX Then
                strTemp = Mid$(strLine, lngX + 1)
                strTemp = mcFile.RetOnlyFilename(strTemp)
                strLine = Mid$(strLine, 1, lngX) & mstrPathNames(1, 5) & strTemp
             Else
                strTemp = Mid$(strLine, 7)
                strTemp = mcFile.RetOnlyFilename(strTemp)
                strLine = Mid$(strLine, 1, 6) & mstrPathNames(1, 5) & strTemp
            End If
         Case "Form=" 'Form=
            strTemp = Mid$(strLine, 6)
            strTemp = mcFile.RetOnlyFilename(strTemp)
            strLine = "Form=" & mstrPathNames(1, 6) & strTemp
         Case "Modul" 'Module=xxxxx; xxxxx.bas
            lngX = InStr(strLine, " ")
            If lngX Then
                strTemp = Mid$(strLine, lngX + 1)
                strTemp = mcFile.RetOnlyFilename(strTemp)
                strLine = Mid$(strLine, 1, lngX) & mstrPathNames(1, 7) & strTemp
             Else
                strTemp = Mid$(strLine, 8)
                strTemp = mcFile.RetOnlyFilename(strTemp)
                strLine = Mid$(strLine, 1, 7) & mstrPathNames(1, 7) & strTemp
            End If
         Case "Relat" 'RelatedDoc=
            strTemp = Mid$(strLine, 12)
            strTemp = mcFile.RetOnlyFilename(strTemp)
            strLine = "RelatedDoc=" & mstrPathNames(1, 8) & strTemp
        End Select
        Print #lngNFN, strLine
    Loop

    Close #lngFN
    Close #lngNFN

    '/* Rename new project name to old project name if source and destination folders are the same
    If blnRefresh Then
        mcFile.DeleteFile strProjectPath & strProjectName
        mcFile.RenameFile strProjectPath & strNewProjectName, strProjectPath & strProjectName
    End If

    CreateNewProjectFile = True

Exit_Here:

Exit Function


Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "frmMain", "CreateNewProjectFile"
    Err.Clear
    Resume Exit_Here

End Function

Private Function CreateNewGroupProjectFile() As Boolean

  Dim lngNFN            As Long
  Dim lngFN             As Long
  Dim blnRefresh        As Boolean
  Dim strTemp           As String
  Dim strProjectPath    As String
  Dim strProjectName    As String
  Dim strNewProjectName As String
  Dim strNewProjectPath As String
  Dim strLine           As String
  Dim strPathPrefix     As String

    On Error GoTo Err_Proc
    
    mcFile.CreateDir mstrNewProjectPath
    
    strProjectPath = mcFile.RetOnlyPath(mstrProjectFileName)
    strProjectName = mcFile.RetOnlyFilename(mstrProjectFileName)

    strNewProjectPath = mstrNewProjectPath
    strNewProjectName = strProjectName

    '/* Check for same folder
    If mcFile.RetShortPathName(strProjectPath) = mcFile.RetShortPathName(strNewProjectPath) Then
        If optMove.Value Then mstrActiveMessage = m_cWarnMessage2
        If frmMsgBox.SMessageModal(C_WarnMessage & mstrActiveMessage & "Do you wish to continue?", vbQuestion + vbYesNo, _
            "WARNING", , , , Me.Width, , Me) = vbNo Then
            
            CreateNewGroupProjectFile = False
            GoTo Exit_Here
        End If
        '/* Make new file name
        strTemp = strNewProjectName
        strTemp = Mid$(strNewProjectName, 1, Len(strNewProjectName) - 4)
        If Len(strNewProjectName) > 14 Then
            If Val(right$(strNewProjectName, 14)) > 0 Then
                strTemp = left$(strTemp, Len(strTemp) - 14)
                strNewProjectName = strTemp & Format$(Now, "yyyymmddhhmmss") & Mid$(strNewProjectName, Len(strNewProjectName) - 3)
            Else
                strNewProjectName = strTemp & Format$(Now, "yyyymmddhhmmss") & Mid$(strNewProjectName, Len(strNewProjectName) - 3)
            End If
        Else
            strNewProjectName = strTemp & Format$(Now, "yyyymmddhhmmss") & Mid$(strNewProjectName, Len(strNewProjectName) - 3)
        End If
        
        cmdAbort.Enabled = False
        blnRefresh = True
        mblnReLoadProject = True
    
    End If

    '/* Begin building new project file
    lngFN = FreeFile
    Open strProjectPath & strProjectName For Input As #lngFN

    lngNFN = FreeFile
    Open strNewProjectPath & strNewProjectName For Output As #lngNFN

    Do
        Line Input #lngFN, strLine

        Select Case left$(strLine, 7)
         Case "Startup"
            strPathPrefix = Mid$(strLine, 1, 15) & mcFile.RetOnlyFilename(Mid$(strLine, 16))
            strPathPrefix = left$(strPathPrefix, Len(strPathPrefix) - 4) & "\"
            strLine = strPathPrefix & mcFile.RetOnlyFilename(Mid$(strLine, 16))

         Case "Project"
            strPathPrefix = Mid$(strLine, 1, 8) & mcFile.RetOnlyFilename(Mid$(strLine, 9))
            strPathPrefix = left$(strPathPrefix, Len(strPathPrefix) - 4) & "\"
            strLine = strPathPrefix & mcFile.RetOnlyFilename(Mid$(strLine, 9))

        End Select

        Print #lngNFN, strLine

    Loop Until EOF(lngFN)
    Close #lngFN
    Close #lngNFN
    CreateNewGroupProjectFile = True

    '/* Copy Support Files in root directory
    Call CopySupportFiles(strProjectPath, mstrNewProjectPath, strPathPrefix)
    
    If Not mblnQuitCommand Then
        '/* Rename new project name to old project name if source and destination folders are the same
        If blnRefresh Then
            mcFile.DeleteFile strProjectPath & strProjectName
            mcFile.RenameFile strProjectPath & strNewProjectName, strProjectPath & strProjectName
        End If
    End If
    
Exit_Here:

Exit Function


Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "frmMain", "CreateNewGroupProjectFile"
    Err.Clear
    CreateNewGroupProjectFile = False
    Resume Exit_Here

End Function


Private Sub cmdQuit_Click()

    Unload frmMain

End Sub

Private Sub cmdAbout_Click()

    picOptions.Visible = False
    frmAbout.Show , Me

End Sub

Private Sub cmdCLose_Click()

    picOptions.Visible = False

End Sub

Private Sub cmdFindFolder_Click()

  Dim strPath As String

    On Error GoTo Err_Proc
    
    picOptions.Visible = False
    
    txtPath.Text = Trim$(txtPath.Text)
    If txtPath.Text > vbNullString Then
        strPath = txtPath.Text
     Else
        strPath = vbNullString
    End If

    '/* Remove if running Win 95 or NT
    strPath = BrowseForFolder(Me, lblMisc.Caption, strPath)

'''    '/* Win 95 and NT users only
'''    With frmBrowseDir
'''        .FolderName = strPath
'''        .Show vbModal
'''        strPath = .FolderName
'''    End With
'''    Unload frmBrowseDir

    If strPath = vbNullString Then Exit Sub

    If right$(strPath, 1) <> "\" Then strPath = strPath & "\"
    txtPath.Text = strPath
    txtPath.ToolTipText = strPath
    txtPath.SelStart = Len(strPath)

Exit_Here:

Exit Sub


Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "frmMain", "cmdFindFolder_Click"
    Err.Clear
    Resume Exit_Here

End Sub

Private Sub cmdOptions_Click()

    picOptions.Visible = Not picOptions.Visible

End Sub

Private Sub Form_Initialize()
    '/* Used for Manifest files (Win XP style controls)
    Call InitCommonControls
End Sub

Private Sub Form_Load()

    On Error GoTo Err_Proc

    Set mcScreen = New clsScreenSize
    mcScreen.vFitScreen Me
    Set mcScreen = Nothing
    
    mstrPathNames(0, 1) = "PropertyPage\"
    mstrPathNames(0, 2) = "UserControl\"
    mstrPathNames(0, 3) = "RelatedDoc\"
    mstrPathNames(0, 4) = "Designer\"
    mstrPathNames(0, 5) = "Class\"
    mstrPathNames(0, 6) = "Form\"
    mstrPathNames(0, 7) = "Module\"
    mstrPathNames(0, 8) = "RelatedDoc\"
    mstrPathNames(0, 9) = "UserDocument\"
    mstrPathNames(0, 10) = "Library\"
    mstrPathNames(0, 11) = "Documentation\"
    
    cmdAbort.Visible = False
    

    Set mcFile = New clsFileUtilities

    '/* Define progress bar parameters
    Set mcProgBar = New clsProgressBar
    With mcProgBar
        .PicBox = picProg
        .Style = pbSolid2Color
    End With
   
    '/* Load last menu picks
    chkSupportDir.Value = CInt(GetSetting(App.Title, "Options", "SupportDir", "1"))
    chkSupportFile.Value = CInt(GetSetting(App.Title, "Options", "SupportFile", "1"))
    chkOrganize.Value = CInt(GetSetting(App.Title, "Options", "Organize", "1"))
    chkDeleteEmpty.Value = CInt(GetSetting(App.Title, "Options", "DeleteEmpty", "1"))
    ''' optMove.Value = CBool(GetSetting(App.Title, "Options", "Move", "True"))
    
    '/* If menu is not default then show
    picOptions.Visible = CBool(chkSupportDir.Value = vbUnchecked _
                                Or chkSupportFile.Value = vbUnchecked _
                                Or chkOrganize.Value = vbUnchecked _
                                Or chkDeleteEmpty.Value = vbUnchecked _
                                Or optMove.Value)

    Call OpenProjectFile

Exit_Here:
Exit Sub


Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "frmMain", "Form_Load"
    Err.Clear
    Resume Exit_Here

End Sub

Private Sub CopyProjectFiles(ByVal pIndex As Long)

  Dim lngI                As Long
  Dim lngN                As Long
  Dim blnMatchFound       As Boolean
  Dim strOldFileName      As String
  Dim strNewFileName      As String
  Dim astrSubDirList()    As String
  Dim ablnDirCrtd(1 To 9) As Boolean
  Dim strProjectPath      As String
  Dim strNewProjectPath   As String
  Dim strPathPrefix       As String


    On Error GoTo Err_Proc

    If LenB(mstrProjectGroups(pIndex)) Then
        strProjectPath = mstrProjectGroups(pIndex)
    Else
        strProjectPath = mstrProjectFileName
    End If
    If strProjectPath = vbNullString Then Exit Sub

    strProjectPath = mcFile.RetOnlyPath(strProjectPath)
    strNewProjectPath = mstrNewProjectPath
    
    If LenB(mstrProjectGroups(pIndex)) Then
        strPathPrefix = mcFile.RetOnlyFilename(mstrProjectGroups(pIndex))
        strPathPrefix = left$(strPathPrefix, Len(strPathPrefix) - 4) & "\"
        strNewProjectPath = strNewProjectPath & strPathPrefix
    End If

    mcProgBar.Max = mcProgList(pIndex).ListCount

    '/* Copy/Move Project Files
    For lngI = 0 To mcProgList(pIndex).ListCount - 1

        strOldFileName = mcProgList(pIndex).List(lngI)
        strNewFileName = mcFile.RetOnlyFilename(strOldFileName)

        Select Case mcProgList(pIndex).ItemData(lngI)
         Case 1 '/* PropertyPage (pag, pgx)
            If Not ablnDirCrtd(1) Then mcFile.CreateDir strNewProjectPath & mstrPathNames(1, 1)
            ablnDirCrtd(1) = True
            CopyMove strOldFileName, strNewProjectPath & mstrPathNames(1, 1) & strNewFileName
            strOldFileName = left$(strOldFileName, Len(strOldFileName) - 3) & "pgx"
            If Dir$(strOldFileName) > vbNullString Then
                CopyMove strOldFileName, strNewProjectPath & mstrPathNames(1, 2) & left$(strNewFileName, Len(strNewFileName) - 3) & "pgx"
            End If

         Case 2 '/* UserControl (ctl, ctx)
            If Not ablnDirCrtd(2) Then mcFile.CreateDir strNewProjectPath & mstrPathNames(1, 2)
            ablnDirCrtd(2) = True
            CopyMove strOldFileName, strNewProjectPath & mstrPathNames(1, 2) & strNewFileName
            strOldFileName = left$(strOldFileName, Len(strOldFileName) - 3) & "ctx"
            If Dir$(strOldFileName) > vbNullString Then
                CopyMove strOldFileName, strNewProjectPath & mstrPathNames(1, 2) & left$(strNewFileName, Len(strNewFileName) - 3) & "ctx"
            End If

         Case 3 '/* ResFile32 (res)
            If Not ablnDirCrtd(3) Then mcFile.CreateDir strNewProjectPath & mstrPathNames(1, 3)
            ablnDirCrtd(3) = True
            CopyMove strOldFileName, strNewProjectPath & mstrPathNames(1, 3) & strNewFileName

         Case 4 '/* Designer (dsr, dca)
            If Not ablnDirCrtd(4) Then mcFile.CreateDir strNewProjectPath & mstrPathNames(1, 4)
            ablnDirCrtd(4) = True
            CopyMove strOldFileName, strNewProjectPath & mstrPathNames(1, 4) & strNewFileName
            strOldFileName = left$(strOldFileName, Len(strOldFileName) - 3) & "dca"
            If Dir$(strOldFileName) > vbNullString Then
                CopyMove strOldFileName, strNewProjectPath & mstrPathNames(1, 4) & left$(strNewFileName, Len(strNewFileName) - 3) & "dca"
            End If

         Case 5 '/* Class (cls)
            If Not ablnDirCrtd(5) Then mcFile.CreateDir strNewProjectPath & mstrPathNames(1, 5)
            ablnDirCrtd(5) = True
            CopyMove strOldFileName, strNewProjectPath & mstrPathNames(1, 5) & strNewFileName

         Case 6 '/* Form (frm, frx)
            If Not ablnDirCrtd(6) Then mcFile.CreateDir strNewProjectPath & mstrPathNames(1, 6)
            ablnDirCrtd(6) = True
            CopyMove strOldFileName, strNewProjectPath & mstrPathNames(1, 6) & strNewFileName
            strOldFileName = left$(strOldFileName, Len(strOldFileName) - 3) & "frx"
            If Dir$(strOldFileName) > vbNullString Then
                CopyMove strOldFileName, strNewProjectPath & mstrPathNames(1, 6) & left$(strNewFileName, Len(strNewFileName) - 3) & "frx"
            End If

         Case 7 '/* Module (bas, gbl)
            If Not ablnDirCrtd(7) Then mcFile.CreateDir strNewProjectPath & mstrPathNames(1, 7)
            ablnDirCrtd(7) = True
            CopyMove strOldFileName, strNewProjectPath & mstrPathNames(1, 7) & strNewFileName

         Case 8 '/* RelatedDoc (?)
            If Not ablnDirCrtd(8) Then mcFile.CreateDir strNewProjectPath & mstrPathNames(1, 8)
            ablnDirCrtd(8) = True
            CopyMove strOldFileName, strNewProjectPath & mstrPathNames(1, 8) & strNewFileName
         
         Case 9 '/* UserDocument (dob, dox)
            If Not ablnDirCrtd(9) Then mcFile.CreateDir strNewProjectPath & mstrPathNames(1, 9)
            ablnDirCrtd(9) = True
            CopyMove strOldFileName, strNewProjectPath & mstrPathNames(1, 9) & strNewFileName
            strOldFileName = left$(strOldFileName, Len(strOldFileName) - 3) & "dox"
            If Dir$(strOldFileName) > vbNullString Then
                CopyMove strOldFileName, strNewProjectPath & mstrPathNames(1, 9) & left$(strNewFileName, Len(strNewFileName) - 3) & "dox"
            End If

        End Select

        '/* Update Status Bar
        mcProgBar.Value = lngI + 1
        lblPrinting.Caption = strPathPrefix & strNewFileName

        DoEvents
        If mblnQuitCommand Then Exit For
    Next lngI
    If mblnQuitCommand Then GoTo Exit_Here
    
    
    '/* Copy Support Files in root folder
    Call CopySupportFiles(strProjectPath, strNewProjectPath, strPathPrefix)
    If mblnQuitCommand Then GoTo Exit_Here
    
    
    '/* Copy Support sub-folder(s)
    If chkSupportDir.Value Then
        If mcFile.RetSubDirList(strProjectPath, astrSubDirList) Then
            mcProgBar.Max = UBound(astrSubDirList) + 1
            For lngI = 0 To UBound(astrSubDirList)
                lblPrinting.Caption = strPathPrefix & astrSubDirList(lngI)
                mcProgBar.Value = lngI + 1
                blnMatchFound = False
                For lngN = 1 To 9
                    '/* Don't copy sub-directories that were created by this utility
                    If LCase$(left$(mstrPathNames(0, lngN), Len(mstrPathNames(0, lngN)) - 1)) = LCase$(astrSubDirList(lngI)) Then
                        blnMatchFound = True
                        Exit For
                    End If
                    DoEvents
                    If mblnQuitCommand Then Exit For
                Next lngN
                If Not blnMatchFound Then
                    If LCase$(strProjectPath & astrSubDirList(lngI)) <> LCase$(strNewProjectPath) Then
                        If optMove.Value Then
                            mcFile.MoveDir strProjectPath & astrSubDirList(lngI), strNewProjectPath
                        Else
                            mcFile.CopyDir strProjectPath & astrSubDirList(lngI), strNewProjectPath
                        End If
                    End If
                End If
                DoEvents
                If mblnQuitCommand Then Exit For
            Next lngI
        End If
    End If
    If mblnQuitCommand Then GoTo Exit_Here
    
    
    '/* Delete empty directories
    If chkDeleteEmpty.Value Then
        If mcFile.RetSubDirList(strProjectPath, astrSubDirList) Then
            mcProgBar.Max = UBound(astrSubDirList) + 1
            For lngI = 0 To UBound(astrSubDirList)
                lblPrinting.Caption = strPathPrefix & astrSubDirList(lngI)
                mcProgBar.Value = lngI + 1
                If LenB(Dir$(strProjectPath & astrSubDirList(lngI) & "\*.*")) = 0 Then
                    mcFile.DeleteDir strProjectPath & astrSubDirList(lngI)
                End If
                DoEvents
                If mblnQuitCommand Then Exit For
            Next lngI
        End If
    End If


Exit_Here:
    lblPrinting.Caption = vbNullString
    Erase ablnDirCrtd
    Erase astrSubDirList

Exit Sub


Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "frmMain", "CopyProjectFiles"
    Err.Clear
    Resume Exit_Here

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    picOptions.Visible = False
End Sub

Private Sub Form_Resize()

    On Error Resume Next

    If Me.Width < 6200 Then Me.Width = 6200
    If Me.Height < 5000 Then Me.Height = 5000
    
    ProjectList.Height = Me.ScaleHeight - ProjectList.top - picStatus.Height - 50
    ProjectList.Width = Me.ScaleWidth

    picProg.left = 10
    picProg.Width = Me.ScaleWidth - 20

    lblPrinting.left = picProg.left
    lblPrinting.Width = picProg.Width
    
    cmdFindFolder.left = Me.ScaleWidth - cmdFindFolder.Width - 100
    txtPath.Width = cmdFindFolder.left - txtPath.left - 100
    lblProject.Width = Me.ScaleWidth
    Line1(0).X2 = Me.ScaleWidth
    Line1(1).X2 = Me.ScaleWidth
    
    With picOptions
        .DrawWidth = 4
        .ForeColor = &HFFCEB9
        .FillStyle = vbFSTransparent
    End With
    picOptions.Line (0, 0)-(picOptions.ScaleWidth, picOptions.ScaleHeight), , B
    
    txtPath.SelStart = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Dim lngI As Long

    On Error Resume Next

    '/* Save last menu picks
    SaveSetting App.Title, "Options", "SupportDir", CStr(chkSupportDir.Value)
    SaveSetting App.Title, "Options", "SupportFile", CStr(chkSupportFile.Value)
    SaveSetting App.Title, "Options", "Organize", CStr(chkOrganize.Value)
    SaveSetting App.Title, "Options", "DeleteEmpty", CStr(chkDeleteEmpty.Value)
    ''' SaveSetting App.Title, "Options", "Move", CStr(optMove.Value)
    
    '/* Clean up
    For lngI = 0 To UBound(mstrProjectGroups)
        Set mcProgList(lngI) = Nothing
    Next lngI
    
    Erase mstrProjectGroups
    Erase mstrPathNames
    
    Set mcProgBar = Nothing
    Set mcFile = Nothing
    Set frmMain = Nothing
    
    '/* End

End Sub

Private Sub RefreshProject(ByVal vlngIndex As Long)

  Dim lngFN               As Long
  Dim strLine             As String
  Dim lngX                As Long
  Dim strProjectName      As String
  Dim strProjectPath      As String
  Dim blnAutoIncrementVer As Boolean
  Dim strTemp             As String


    On Error GoTo Err_Proc

    If LenB(mstrProjectGroups(vlngIndex)) Then
        strProjectName = mstrProjectGroups(vlngIndex)
     Else
        strProjectName = mstrProjectFileName
    End If

    If strProjectName = vbNullString Then Exit Sub

    strProjectPath = mcFile.RetOnlyPath(strProjectName)
    strProjectName = mcFile.RetOnlyFilename(strProjectName)
    
    If vlngIndex = 0 Then ProjectList.AddItem String$(500, "*")

    lngFN = FreeFile
    Open strProjectPath & strProjectName For Input As #lngFN

    Do
        Line Input #lngFN, strLine

        If Not mblnIsVBG Then '/* Single Project File
            Select Case left$(strLine, 5)
             Case "Prope" 'PropertyPage=
                strTemp = Mid$(strLine, 14)
                ProjectList.AddItem "PropertyPage - " & strTemp
                If InStr(strTemp, ":") = 0 Then
                    strTemp = strProjectPath & strTemp
                End If
                mcProgList(vlngIndex).ItemDataSet = 1
                mcProgList(vlngIndex).AddItem strTemp
             Case "UserC" 'UserControl=
                strTemp = Mid$(strLine, 13)
                ProjectList.AddItem "UserControl - " & strTemp
                If InStr(strTemp, ":") = 0 Then
                    strTemp = strProjectPath & strTemp
                End If
                mcProgList(vlngIndex).ItemDataSet = 2
                mcProgList(vlngIndex).AddItem strTemp
             Case "UserD" 'UserDocument=
                strTemp = Mid$(strLine, 14)
                ProjectList.AddItem "UserDocument - " & strTemp
                If InStr(strTemp, ":") = 0 Then
                    strTemp = strProjectPath & strTemp
                End If
                mcProgList(vlngIndex).ItemDataSet = 9
                mcProgList(vlngIndex).AddItem strTemp
             Case "ResFi" 'ResFile32="XXXX.res"
                strTemp = mcFile.StripQuotes(Mid$(strLine, 11))
                ProjectList.AddItem "ResFile32    - " & strTemp
                If InStr(strTemp, ":") = 0 Then
                    strTemp = strProjectPath & strTemp
                End If
                mcProgList(vlngIndex).ItemDataSet = 3
                mcProgList(vlngIndex).AddItem strTemp
             Case "Desig" 'Designer=
                strTemp = Mid$(strLine, 10)
                ProjectList.AddItem "Designer      - " & strTemp
                If InStr(strTemp, ":") = 0 Then
                    strTemp = strProjectPath & strTemp
                End If
                mcProgList(vlngIndex).ItemDataSet = 4
                mcProgList(vlngIndex).AddItem strTemp
             Case "Class" 'Class=xxxxx; xxxxx.cls
                lngX = InStr(strLine, " ")
                If lngX Then
                    strTemp = Mid$(strLine, lngX + 1)
                Else
                    strTemp = Mid$(strLine, 7)
                End If
                ProjectList.AddItem "Class           - " & strTemp
                If InStr(strTemp, ":") = 0 Then
                    strTemp = strProjectPath & strTemp
                End If
                mcProgList(vlngIndex).ItemDataSet = 5
                mcProgList(vlngIndex).AddItem strTemp
             Case "Form=" 'Form=
                strTemp = Mid$(strLine, 6)
                ProjectList.AddItem "Form            - " & strTemp
                If InStr(strTemp, ":") = 0 Then
                    strTemp = strProjectPath & strTemp
                End If
                mcProgList(vlngIndex).ItemDataSet = 6
                mcProgList(vlngIndex).AddItem strTemp
             Case "Modul" 'Module=xxxxx; xxxxx.bas
                lngX = InStr(strLine, " ")
                If lngX Then
                    strTemp = Mid$(strLine, lngX + 1)
                 Else
                    strTemp = Mid$(strLine, 8)
                End If
                ProjectList.AddItem "Module        - " & strTemp
                If InStr(strTemp, ":") = 0 Then
                    strTemp = strProjectPath & strTemp
                End If
                mcProgList(vlngIndex).ItemDataSet = 7
                mcProgList(vlngIndex).AddItem strTemp
             Case "Relat" 'RelatedDoc=
                strTemp = Mid$(strLine, 12)
                ProjectList.AddItem "RelatedDoc - " & strTemp
                If InStr(strTemp, ":") = 0 Then
                    strTemp = strProjectPath & strTemp
                End If
                mcProgList(vlngIndex).ItemDataSet = 8
                mcProgList(vlngIndex).AddItem strTemp
                
             Case "Refer" 'Reference=
                For lngX = Len(strLine) To 1 Step -1
                    If Mid$(strLine, lngX, 1) = "#" Then
                        ProjectList.AddItem " Reference - " & Mid$(strLine, lngX + 1)
                        Exit For
                    End If
                Next lngX
             Case "Objec" 'Object=
                lngX = InStr(strLine, ";")
                ProjectList.AddItem "       Object - " & Mid$(strLine, lngX + 2)
             Case "Major" 'MajorVer=
                mudtProjectInfo.MajorVer = Mid$(strLine, 10)
             Case "Minor" 'MinorVer=
                mudtProjectInfo.MinorVer = Mid$(strLine, 10)
             Case "Revis" 'RevisionVer=
                mudtProjectInfo.RevisionVer = Mid$(strLine, 13)
                mudtProjectInfo.RevisionVer = Format$(Val(mudtProjectInfo.RevisionVer), "00")
             Case "Name=" 'Name=
                mudtProjectInfo.PName = mcFile.StripQuotes(Mid$(strLine, 6))
             Case "AutoI" 'AutoIncrementVer=
                blnAutoIncrementVer = Not CBool(right$(strLine, 1))
            End Select
         
         Else '/* Group Project File
            
            Select Case left$(strLine, 5)
             Case "Prope" 'PropertyPage=
                strTemp = Mid$(strLine, 14)
                If InStr(strTemp, ":") = 0 Then
                    strTemp = strProjectPath & strTemp
                End If
                mcProgList(vlngIndex).ItemDataSet = 1
                mcProgList(vlngIndex).AddItem strTemp
             Case "UserC" 'UserControl=
                strTemp = Mid$(strLine, 13)
                If InStr(strTemp, ":") = 0 Then
                    strTemp = strProjectPath & strTemp
                End If
                mcProgList(vlngIndex).ItemDataSet = 2
                mcProgList(vlngIndex).AddItem strTemp
             Case "UserD" 'UserDocument=
                strTemp = Mid$(strLine, 14)
                If InStr(strTemp, ":") = 0 Then
                    strTemp = strProjectPath & strTemp
                End If
                mcProgList(vlngIndex).ItemDataSet = 9
                mcProgList(vlngIndex).AddItem strTemp
             Case "ResFi" 'ResFile32="XXXX.res"
                strTemp = mcFile.StripQuotes(Mid$(strLine, 11))
                If InStr(strTemp, ":") = 0 Then
                    strTemp = strProjectPath & strTemp
                End If
                mcProgList(vlngIndex).ItemDataSet = 3
                mcProgList(vlngIndex).AddItem strTemp
             Case "Desig" 'Designer=
                strTemp = Mid$(strLine, 10)
                If InStr(strTemp, ":") = 0 Then
                    strTemp = strProjectPath & strTemp
                End If
                mcProgList(vlngIndex).ItemDataSet = 4
                mcProgList(vlngIndex).AddItem strTemp
             Case "Class" 'Class=xxxxx; xxxxx.cls
                lngX = InStr(strLine, " ")
                If lngX Then
                    strTemp = Mid$(strLine, lngX + 1)
                Else
                    strTemp = Mid$(strLine, 7)
                End If
                If InStr(strTemp, ":") = 0 Then
                    strTemp = strProjectPath & strTemp
                End If
                mcProgList(vlngIndex).ItemDataSet = 5
                mcProgList(vlngIndex).AddItem strTemp
             Case "Form=" 'Form=
                strTemp = Mid$(strLine, 6)
                If InStr(strTemp, ":") = 0 Then
                    strTemp = strProjectPath & strTemp
                End If
                mcProgList(vlngIndex).ItemDataSet = 6
                mcProgList(vlngIndex).AddItem strTemp
             Case "Modul" 'Module=xxxxx; xxxxx.bas
                lngX = InStr(strLine, " ")
                If lngX Then
                    strTemp = Mid$(strLine, lngX + 1)
                Else
                    strTemp = Mid$(strLine, 8)
                End If
                If InStr(strTemp, ":") = 0 Then
                    strTemp = strProjectPath & strTemp
                End If
                mcProgList(vlngIndex).ItemDataSet = 7
                mcProgList(vlngIndex).AddItem strTemp
             
             Case "Relat" 'RelatedDoc=
                strTemp = Mid$(strLine, 12)
                If InStr(strTemp, ":") = 0 Then
                    strTemp = strProjectPath & strTemp
                End If
                mcProgList(vlngIndex).ItemDataSet = 8
                mcProgList(vlngIndex).AddItem strTemp

             Case "Refer" 'Reference=
                For lngX = Len(strLine) To 1 Step -1
                    If Mid$(strLine, lngX, 1) = "#" Then
                        Call AddItemNoDup(" Reference - " & Mid$(strLine, lngX + 1))
                        Exit For
                    End If
                Next lngX
             Case "Objec" 'Object=
                lngX = InStr(strLine, ";")
                Call AddItemNoDup("       Object - " & Mid$(strLine, lngX + 2))
             Case "Major" 'MajorVer=
                mudtProjectInfo.MajorVer = Mid$(strLine, 10)
             Case "Minor" 'MinorVer=
                mudtProjectInfo.MinorVer = Mid$(strLine, 10)
             Case "Revis" 'RevisionVer=
                mudtProjectInfo.RevisionVer = Mid$(strLine, 13)
                mudtProjectInfo.RevisionVer = Format$(Val(mudtProjectInfo.RevisionVer), "00")
             Case "Name=" 'Name=
                mudtProjectInfo.PName = mcFile.StripQuotes(Mid$(strLine, 6))
             Case "AutoI" 'AutoIncrementVer=
                blnAutoIncrementVer = Not CBool(right$(strLine, 1))
            End Select
            
        End If
        
    Loop Until EOF(lngFN)
    
    Close #lngFN

    '/* If you are automatically incrementing the version, then assume it was compiled
    '/* and decrement the version number to match the exe.
    If blnAutoIncrementVer Then
        If Val(mudtProjectInfo.RevisionVer) > 0 Then
            mudtProjectInfo.RevisionVer = Format$(Val(mudtProjectInfo.RevisionVer) - 1, "00")
        End If
    End If

    If vlngIndex = 0 Then
        If mblnIsVBG Then
            lblProject.Caption = "Group Project: "
        Else
            lblProject.Caption = "Project: "
        End If
        lblProject.Caption = lblProject.Caption & mudtProjectInfo.PName & "  v" & mudtProjectInfo.MajorVer & "." & mudtProjectInfo.MinorVer & "." & mudtProjectInfo.RevisionVer
    End If

Exit_Here:

Exit Sub


Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "frmMain", "RefreshProject"
    Err.Clear
    Resume Exit_Here

End Sub

Private Sub OpenProjectFile()

  Dim lngI As Long

    On Error GoTo Err_Proc
    Screen.MousePointer = vbHourglass
    
    mstrProjectFileName = VBInstance.VBProjects.FileName
    If LenB(Dir$(mstrProjectFileName)) = 0 Then
        mstrProjectFileName = VBInstance.ActiveVBProject.FileName
    End If


    '/* vb project - add all relevant files
    ProjectList.Clear
    ReDim mstrProjectGroups(0)
    ReDim mcProgList(0)

    If mcFile.GetExtensionName(mstrProjectFileName) = "vbg" Then
        mblnIsVBG = True
        Call GetGroupProjects
        For lngI = 0 To UBound(mstrProjectGroups)
            Call RefreshProject(lngI)
        Next lngI
     Else
        mblnIsVBG = False
        mstrProjectGroups(0) = vbNullString
        Set mcProgList(0) = New clsVirtualListbox
        Call RefreshProject(0)
    End If

    mstrProjectFilePath = mcFile.RetOnlyPath(mstrProjectFileName)


Exit_Here:
    Screen.MousePointer = vbDefault
Exit Sub


Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "frmMain", "OpenProjectFile"
    Err.Clear
    Resume Exit_Here

End Sub

Private Sub GetGroupProjects()

  Dim lngFN          As Long
  Dim lngI           As Long
  Dim strLine        As String
  Dim strProjectPath As String

    strProjectPath = mcFile.RetOnlyPath(mstrProjectFileName)
    
    '/* Open group project file
    lngFN = FreeFile
    Open mstrProjectFileName For Input As #lngFN
    Do
        Line Input #lngFN, strLine

        Select Case left$(strLine, 7)
         Case "Startup"
            mstrProjectGroups(lngI) = Mid$(strLine, 16)
            If InStr(mstrProjectGroups(lngI), ":") = 0 Then
                mstrProjectGroups(lngI) = strProjectPath & mstrProjectGroups(lngI)
            End If
            lngI = UBound(mstrProjectGroups) + 1
            ReDim Preserve mstrProjectGroups(lngI)
            ProjectList.AddItem mcFile.RetOnlyFilename(Mid$(strLine, 16))

         Case "Project"
            mstrProjectGroups(lngI) = Mid$(strLine, 9)
            If InStr(mstrProjectGroups(lngI), ":") = 0 Then
                mstrProjectGroups(lngI) = strProjectPath & mstrProjectGroups(lngI)
            End If
            lngI = UBound(mstrProjectGroups) + 1
            ReDim Preserve mstrProjectGroups(lngI)
            ProjectList.AddItem mcFile.RetOnlyFilename(Mid$(strLine, 9))

        End Select

    Loop Until EOF(lngFN)
    Close #lngFN

    lngI = lngI - 1
    ReDim Preserve mstrProjectGroups(lngI)

    ReDim mcProgList(lngI)
    For lngI = 0 To UBound(mstrProjectGroups)
        Set mcProgList(lngI) = New clsVirtualListbox
    Next lngI

End Sub

Private Sub lblMisc_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    picOptions.Visible = False
End Sub

Private Sub lblPrinting_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    picOptions.Visible = False
End Sub

Private Sub lblProject_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    picOptions.Visible = False
End Sub

Private Sub optCopy_Click()
    
    If optCopy.Value Then cmdStart.Caption = "Start Copy"

End Sub

Private Sub optMove_Click()
    
    If optMove.Value Then cmdStart.Caption = "Start Move"
    
End Sub


Private Sub picStatus_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    picOptions.Visible = False
End Sub

Private Sub picToolBar_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    picOptions.Visible = False
End Sub

Private Sub ProjectList_KeyUp(KeyCode As Integer, Shift As Integer)
    If ProjectList.ListIndex > -1 Then
        ProjectList.Selected(ProjectList.ListIndex) = False
    End If
End Sub

Private Sub ProjectList_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    picOptions.Visible = False
    If ProjectList.ListIndex > -1 Then
        ProjectList.Selected(ProjectList.ListIndex) = False
    End If
End Sub

Private Sub txtPath_Change()
    
    If LCase$(mstrProjectFilePath) = LCase$(IIf(right$(txtPath.Text, 1) = "\", txtPath.Text, txtPath.Text & "\")) Then
        optMove.Value = True
    End If
    
End Sub

Private Sub txtPath_KeyPress(KeyAscii As Integer)
    '/* don't allow ",*,/,<,>,?,|
    Select Case KeyAscii
     Case Is < 32, Is > 126, 34, 42, 47, 60, 62, 63, 124
        KeyAscii = 0
     Case Else
    End Select
End Sub

Private Sub txtPath_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    picOptions.Visible = False
End Sub

Private Sub AddItemNoDup(ByVal sAddItem As String)
  
  Dim lngI As Long
  Dim blnMatchFound As Boolean
    
    '/* Don't add duplicates to list box
    For lngI = 0 To ProjectList.ListCount - 1
        If ProjectList.List(lngI) = sAddItem Then
            blnMatchFound = True
            Exit For
        End If
    Next lngI
    If Not blnMatchFound Then
        ProjectList.AddItem sAddItem
    End If
    
End Sub

Private Sub CopyMove(ByVal vstrSPath As String, ByVal vstrDPath As String)

    If optMove.Value Then
        
        If InStr(vstrSPath, "..\") = 0 And LCase$(left$(vstrSPath, 1)) = LCase$(left$(vstrDPath, 1)) Then
            mcFile.MoveFile vstrSPath, vstrDPath, mblnOnlyNewer, True
            Exit Sub
        End If
        
    End If
    
    mcFile.CopyFile vstrSPath, vstrDPath, mblnOnlyNewer, True
    
End Sub

Private Sub CopySupportFiles(ByVal vstrProjectPath As String, _
                             ByVal vstrNewProjectPath As String, _
                             ByVal vstrPathPrefix As String)
  
  Dim ablnDirCrtd(1 To 2) As Boolean
  Dim lngN                As Long
  Dim astrFileList()      As String
    
    '/* Copy Support Files in root directory
    If chkSupportFile.Value Then
        If mcFile.RetFileList(vstrProjectPath & "*.*", astrFileList) Then
            mcProgBar.Max = UBound(astrFileList)
            For lngN = 0 To UBound(astrFileList)
                Select Case LCase$(right$(astrFileList(lngN), 3))
                 Case "dsr", "dca", "frm", "frx", "cls", "bas", "gbl", "pag", "pgx", _
                        "ctl", "ctx", "res", "vbp", "vbg", "vbw", "dob", "dox"
                    '/* Do nothing
                 Case "bmp", "gif", "jpg", "wmf", "ico", "cur", "wav"
                    If Not ablnDirCrtd(1) Then mcFile.CreateDir vstrNewProjectPath & mstrPathNames(1, 10) '/* Library
                    ablnDirCrtd(1) = True
                    CopyMove vstrProjectPath & astrFileList(lngN), vstrNewProjectPath & mstrPathNames(1, 10) & astrFileList(lngN)
                 Case "txt", "doc", "htm", "rtf"
                    If Not ablnDirCrtd(2) Then mcFile.CreateDir vstrNewProjectPath & mstrPathNames(1, 11) '/* Documentation
                    ablnDirCrtd(2) = True
                    CopyMove vstrProjectPath & astrFileList(lngN), vstrNewProjectPath & mstrPathNames(1, 11) & astrFileList(lngN)
                 Case Else
                    CopyMove vstrProjectPath & astrFileList(lngN), vstrNewProjectPath & astrFileList(lngN)
                End Select
                
                lblPrinting.Caption = vstrPathPrefix & astrFileList(lngN)
                mcProgBar.Value = lngN + 1
                DoEvents
                If mblnQuitCommand Then Exit For
            Next lngN
        End If
    End If

End Sub
