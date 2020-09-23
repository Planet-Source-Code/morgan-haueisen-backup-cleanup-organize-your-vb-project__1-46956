Attribute VB_Name = "modIsProcessRunning"
Option Explicit

Private Const TH32CS_SNAPHEAPLIST As Long = &H1
Private Const TH32CS_SNAPPROCESS  As Long = &H2
Private Const TH32CS_SNAPTHREAD   As Long = &H4
Private Const TH32CS_SNAPMODULE   As Long = &H8
Private Const TH32CS_INHERIT      As Long = &H80000000
Private Const TH32CS_SNAPALL      As Long = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)

Private Const MAX_PATH As Long = 260

Private Type PROCESSENTRY32
    dwSize              As Long
    cntUsage            As Long
    th32ProcessID       As Long
    th32DefaultHeapID   As Long
    th32ModuleID        As Long
    cntThreads          As Long
    th32ParentProcessID As Long
    pcPriClassBase      As Long
    dwFlags             As Long
    szExeFile           As String * MAX_PATH
End Type

Private Declare Function CreateToolhelp32Snapshot Lib "Kernel32" _
    (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "Kernel32" _
    (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "Kernel32" _
    (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "Kernel32" (ByVal hPass As Long)

Public Function IsActive(ByVal vstrFileName As String) As Boolean

  Dim lngSnapShot As Long
  Dim udtProcess  As PROCESSENTRY32
  Dim lngR        As Long

    vstrFileName = UCase$(vstrFileName)

    '/* Takes a snapshot of the processes and the heaps, modules, and threads used by the processes
    lngSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
    '/* set the length of our ProcessEntry-type
    udtProcess.dwSize = Len(udtProcess)
    '/* Retrieve information about the first process encountered in our system snapshot
    lngR = Process32First(lngSnapShot, udtProcess)

    Do While lngR
        If vstrFileName = UCase$(left(udtProcess.szExeFile, IIf(InStr(1, udtProcess.szExeFile, vbNullChar) > 0, InStr(1, udtProcess.szExeFile, vbNullChar) - 1, 0))) Then
            IsActive = True
            Exit Do
        End If
        lngR = Process32Next(lngSnapShot, udtProcess)
    Loop

    '/* close our snapshot handle
    CloseHandle lngSnapShot

End Function

