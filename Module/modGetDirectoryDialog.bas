Attribute VB_Name = "modBrowseDirectorysOnly"
' Browse for a Folder using SHBrowseForFolder API function with a callback function BrowseCallbackProc.
' This Extends the functionality that was given in the
' MSDN Knowledge Base article Q179497 "HOWTO: Select a Directory without the Common Dialog Control".

Option Explicit

Private Const BIF_STATUSTEXT              As Long = &H4&
Private Const BIF_RETURNONLYFSDIRS        As Long = &H1
Private Const BIF_DONTGOBELOWDOMAIN       As Long = &H2
Private Const BIF_EDITBOX                 As Long = &H10
Private Const BIF_NEWDIALOGSTYLE          As Long = &H40

Private Const MAX_PATH                    As Long = &H104
Private Const WM_USER                     As Long = &H400

Private Const BFFM_INITIALIZED            As Long = &H1
Private Const BFFM_SELCHANGED             As Long = &H2
Private Const BFFM_SETSTATUSTEXT          As Long = &H464
Private Const BFFM_SETSELECTION           As Long = &H466

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" _
    (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "Kernel32" Alias "lstrcatA" _
    (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
    hWndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private mstrCurrentDirectory As String    '/* The current directory

Public Function BrowseForFolder(Optional ByRef rOwnerForm As Form, _
                                Optional ByVal vstrTitle As String, _
                                Optional ByVal vstrStartDir As String = vbNullString, _
                                Optional ByVal vblnShowCreateNewFolder As Boolean = True) As String

  Dim lngpIDList    As Long
  Dim strBuffer     As String
  Dim udtBrowseInfo As BrowseInfo
  Dim lnghWnd       As Long

    On Error Resume Next

    lnghWnd = rOwnerForm.lnghWnd
    mstrCurrentDirectory = vstrStartDir & vbNullChar

    With udtBrowseInfo
        .hWndOwner = lnghWnd
        .pIDLRoot = 0&
        .lpszTitle = lstrcat(vstrTitle, vbNullChar)
        If vblnShowCreateNewFolder Then
            .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or BIF_STATUSTEXT Or BIF_NEWDIALOGSTYLE
         Else
            .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or BIF_STATUSTEXT
        End If
        .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)  '/* get address of function.
    End With

    lngpIDList = SHBrowseForFolder(udtBrowseInfo)
    If (lngpIDList) Then
        strBuffer = Space$(MAX_PATH)
        SHGetPathFromIDList lngpIDList, strBuffer
        strBuffer = left(strBuffer, InStr(strBuffer, vbNullChar) - 1)
        BrowseForFolder = strBuffer
     Else
        BrowseForFolder = vbNullString
    End If

End Function

Public Function RetOnlyFilename(ByVal vstrPathFileName As String)

    RetOnlyFilename = right(vstrPathFileName, (Len(vstrPathFileName) - InStrRev(vstrPathFileName, "\", , vbTextCompare)))

End Function

Private Function BrowseCallbackProc(ByVal vlnghWnd As Long, _
                                    ByVal vlngMsg As Long, _
                                    ByVal vlngLP As Long, _
                                    ByVal vlngpData As Long) As Long

  Dim lngpIDList As Long
  Dim strBuffer  As String

    On Local Error Resume Next  '/* Sugested by MS to prevent an error from
    '/* propagating back into the calling process.

    Select Case vlngMsg
     Case BFFM_INITIALIZED
        Call SendMessage(vlnghWnd, BFFM_SETSELECTION, 1, mstrCurrentDirectory)
     Case BFFM_SELCHANGED
        strBuffer = Space$(MAX_PATH)
        If SHGetPathFromIDList(vlngLP, strBuffer) = 1 Then
            Call SendMessage(vlnghWnd, BFFM_SETSTATUSTEXT, 0, strBuffer)
        End If
    End Select

    BrowseCallbackProc = 0

End Function

Private Function GetAddressofFunction(ByRef rlngAddress As Long) As Long

  '/* This function allows you to assign a function pointer to a vaiable.

    GetAddressofFunction = rlngAddress

End Function

