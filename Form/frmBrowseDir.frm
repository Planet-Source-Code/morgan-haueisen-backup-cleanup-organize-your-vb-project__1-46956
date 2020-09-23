VERSION 5.00
Begin VB.Form frmBrowseDir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save To"
   ClientHeight    =   4815
   ClientLeft      =   3120
   ClientTop       =   2160
   ClientWidth     =   6495
   Icon            =   "frmBrowseDir.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin BackUpVBProject.chameleonButton cmdOpen 
      Height          =   375
      Left            =   4575
      TabIndex        =   3
      Top             =   4005
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Ok"
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
      MICON           =   "frmBrowseDir.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.DirListBox Dir1 
      Height          =   3465
      Left            =   60
      TabIndex        =   2
      Top             =   465
      Width           =   6315
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   60
      Width           =   3930
   End
   Begin BackUpVBProject.chameleonButton cmdQuit 
      Height          =   375
      Left            =   4575
      TabIndex        =   4
      Top             =   4410
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cancel"
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
      MICON           =   "frmBrowseDir.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin BackUpVBProject.chameleonButton cmdUpOne 
      Height          =   345
      Left            =   5550
      TabIndex        =   5
      Top             =   45
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmBrowseDir.frx":0044
      PICN            =   "frmBrowseDir.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin BackUpVBProject.chameleonButton cmdNewFolder 
      Height          =   345
      Left            =   5955
      TabIndex        =   6
      Top             =   45
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmBrowseDir.frx":0322
      PICN            =   "frmBrowseDir.frx":033E
      UMCOL           =   -1  'True
      SOFT            =   -1  'True
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " Select a Directory: "
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1395
   End
End
Attribute VB_Name = "frmBrowseDir"
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

Public FolderName As String

Private cScreen   As clsScreenSize

'/* Used for Manifest files (Win XP style controls)
Private Declare Sub InitCommonControls Lib "Comctl32.dll" ()

Private Sub cmdNewFolder_Click()

  Dim NewFolderName As String

    On Error GoTo Err_Proc


    NewFolderName = frmMsgBox.SInputBox("Enter Folder Name", , "New Folder")
    NewFolderName = Trim$(NewFolderName)
    If NewFolderName > vbNullString Then
        Call CreateDir(Dir1.Path & "\" & NewFolderName)
        NewFolderName = Dir1.Path & "\" & NewFolderName
        Dir1.Refresh
        Dir1.Path = NewFolderName
    End If

Exit_Here:

Exit Sub


Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "frmBrowseDir", "cmdNewFolder_Click"
    Err.Clear
    Resume Exit_Here

End Sub

Private Function CreateDir(ByVal oPath As String) As Boolean

  '/***********************************************************
  '/* Check to see if the directory exists and make if required
  '/***********************************************************

  Dim start     As Integer
  Dim pos       As Integer
  Dim directory As String

    On Error GoTo errCreation

    oPath = Trim$(oPath)
    '/* if null string why bother....
    If oPath = vbNullString Then Err.Raise vbObjectError + 1
    pos = 0
    If right(oPath, 1) = "\" Then oPath = left(oPath, Len(oPath) - 1)

TryAgain:
    start = pos + 1
    pos = InStr(start, oPath, "\")

    If pos > 0 Then
        '/* not at the last directory in the path string...
        directory = directory & Mid$(oPath, start, pos - start) & "\"
        If InStr(1, Mid$(oPath, start, pos - start), ":") = 0 And Dir(directory, vbDirectory) = vbNullString Then
            MkDir Mid$(directory, 1, Len(directory) - 1)
        End If
        GoTo TryAgain
     Else
        '/* the last or only directory in the path string
        directory = directory & Mid$(oPath, start, Len(oPath) - start + 1)
        MkDir Mid$(directory, 1, Len(directory))
        directory = vbNullString
    End If

    '/* success return true
    On Error GoTo 0
    CreateDir = True

Exit Function


    '/* if it gets here, an exception was thrown
    '/* propogate the error to the calling function
errCreation:
    Err.Clear
    CreateDir = False

End Function

Private Sub cmdOpen_Click()

    FolderName = Dir1.Path
    Me.Hide

End Sub

Private Sub cmdQuit_Click()

    FolderName = vbNullString
    Me.Hide

End Sub

Private Sub cmdUpOne_Click()

    Dir1.Path = Dir1.List(-2)

End Sub

Private Sub Dir1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)

    Dir1.Path = Dir1.List(Dir1.ListIndex)

End Sub

Private Sub Drive1_Change()

    On Error GoTo DriveNotReady

    Dir1.Path = Drive1.Drive

Exit Sub


DriveNotReady:
    frmMsgBox.SMessageModal "Drive: " & Drive1.Drive & vbCrLf & Err.Description, vbExclamation
    Drive1.Drive = CurDir

End Sub

Private Sub Form_Initialize()

  '/* Used for Manifest files (Win XP style controls)
    Call InitCommonControls

End Sub

Private Sub Form_Load()

    Set cScreen = New clsScreenSize
    cScreen.CenterForm Me
    Set cScreen = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmBrowseDir = Nothing

End Sub
