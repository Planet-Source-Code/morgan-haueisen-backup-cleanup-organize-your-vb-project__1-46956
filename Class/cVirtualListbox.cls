VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVirtualListbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mstrItemArray() As String
Private mlngDataArray() As Long

Private mlngItemData    As Long
Private mblnSorted      As Boolean

Public Sub AddItem(ByVal vstrItem As String, _
                   Optional ByVal vlngIndex As Long = -1)

  Dim lngI As Long
  Dim lngN As Long

    On Error GoTo Err_Proc

    '/* The Index requested exceeds the size of the array.
    If vlngIndex >= 0 Then
        If Not IsDim(mstrItemArray) Then GoTo ErrorHandle
        If vlngIndex > UBound(mstrItemArray) Then GoTo ErrorHandle
    End If

    '/* Resize the array.
    If Not IsDim(mstrItemArray) Then
        ReDim mstrItemArray(0)
        ReDim mlngDataArray(0)
     Else
        If mblnSorted Then
            For lngN = 0 To UBound(mstrItemArray)
                If vstrItem < mstrItemArray(lngN) Then
                    vlngIndex = lngN
                    Exit For
                End If
            Next lngN
        End If
        lngI = UBound(mstrItemArray) + 1
        ReDim Preserve mstrItemArray(lngI)
        ReDim Preserve mlngDataArray(lngI)
    End If

    If vlngIndex = -1 Then
        '/* If index is -1 then add the item to the bottom on the array
        mstrItemArray(UBound(mstrItemArray)) = vstrItem
        mlngDataArray(UBound(mstrItemArray)) = mlngItemData
     Else
        '/* If index > 0 then move current value to next position starting at the end
        '/* and working forward.
        For lngI = UBound(mstrItemArray) To vlngIndex + 1 Step -1
            mstrItemArray(lngI) = mstrItemArray(lngI - 1)
        Next lngI
        '/* After moving all the indexes around, add this item to the specified index
        mstrItemArray(vlngIndex) = vstrItem
        mlngDataArray(vlngIndex) = mlngItemData
    End If

Exit_Proc:

Exit Sub


ErrorHandle:
    Err_Handler True, 0, "Subscript Error: You're attempting to add an item in a place that doesn't exist!", "clssItemArray", "AddItem"

Exit Sub


Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "clssItemArray", "AddItem"
    Err.Clear
    Resume Exit_Proc

End Sub

Private Sub Err_Handler(Optional ByVal vblnDisplayError As Boolean = True, _
                        Optional ByVal vstrErrNumber As String = vbNullString, _
                        Optional ByVal vstrErrDescription As String = vbNullString, _
                        Optional ByVal vstrModuleName As String = vbNullString, _
                        Optional ByVal vstrProcName As String = vbNullString)

  Dim strTemp As String
  Dim lngFN   As Long

    '/* Purpose: Error handling - On Error

    '/* Show Error Message
    If vblnDisplayError Then
        strTemp = "Error occured: "
        If Len(vstrErrNumber) > 0 Then strTemp = strTemp & vstrErrNumber & vbNewLine Else strTemp = strTemp & vbNewLine
        If Len(vstrErrDescription) > 0 Then strTemp = strTemp & "Description: " & vstrErrDescription & vbNewLine
        If Len(vstrModuleName) > 0 Then strTemp = strTemp & "Module: " & vstrModuleName & vbNewLine
        If Len(vstrProcName) > 0 Then strTemp = strTemp & "Function: " & vstrProcName
        MsgBox strTemp, vbCritical, App.Title & " - ERROR"
    End If

    '/* Write error log
    lngFN = FreeFile
    Open App.Path & "\ErrorLog.txt" For Append As #lngFN
    Write #lngFN, Now, vstrErrNumber, vstrErrDescription, vstrModuleName, vstrProcName, App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision, Environ$("username"), Environ$("computername")
    Close #lngFN

End Sub

Public Sub RemoveItem(ByVal vlngIndex As Long)

  Dim lngI As Long

    On Error GoTo Err_Proc

    '/* The Index requested exceeds the size of the array.
    If Not IsDim(mstrItemArray) Then GoTo ErrorHandle
    If vlngIndex < 0 Or vlngIndex > UBound(mstrItemArray) Then GoTo ErrorHandle

    For lngI = vlngIndex To UBound(mstrItemArray)
        mstrItemArray(lngI) = mstrItemArray(lngI + 1)
        mlngDataArray(lngI) = mlngDataArray(lngI + 1)
    Next lngI

    ReDim Preserve mstrItemArray(UBound(mstrItemArray) - 1)
    ReDim Preserve mlngDataArray(UBound(mstrItemArray))

Exit_Proc:

Exit Sub


ErrorHandle:
    Err_Handler True, 0, "Subscript Error: You're attempting to remove an item in a place that doesn't exist!", "clssItemArray", "RemoveItem"

Exit Sub


Err_Proc:
    Err_Handler True, Err.Number, Err.Description, "clssItemArray", "RemoveItem"
    Err.Clear
    Resume Exit_Proc

End Sub

Public Sub RemoveDuplicates()

  Dim lngI As Long
  Dim lngY As Long

    '/* No items in array, exit sub.
    If Not IsDim(mstrItemArray) Then Exit Sub
    If UBound(mstrItemArray) < 2 Then Exit Sub

    For lngI = 0 To UBound(mstrItemArray)
        For lngY = 0 To UBound(mstrItemArray)
            If lngY <> lngI Then
                '/* Remove item index if it appears more than once.
                If LCase$(mstrItemArray(lngY)) = LCase$(mstrItemArray(lngI)) Then
                    Call RemoveItem(lngY)
                    lngY = lngY - 1
                End If
            End If
        Next lngY
    Next lngI

End Sub

Public Sub RemoveItemByString(ByVal vstrItem As String)

  Dim lngI As Long

    For lngI = 0 To UBound(mstrItemArray)
        If LCase$(mstrItemArray(lngI)) = LCase$(vstrItem) Then
            Call RemoveItem(lngI)
        End If
    Next lngI

End Sub

Public Sub Clear()

    Erase mstrItemArray
    Erase mlngDataArray

End Sub

Public Property Get ItemData(ByVal vlngIndex As Long) As Long

    mlngItemData = mlngDataArray(vlngIndex)
    ItemData = mlngItemData

End Property

Public Property Let ItemDataSet(ByVal vNewItemValue As Long)

  '/* Set value here before using AddItem

    mlngItemData = vNewItemValue

End Property

Public Property Get ListCount() As Long

    If IsDim(mstrItemArray) Then
        ListCount = UBound(mstrItemArray) + 1
     Else
        ListCount = -1
    End If

End Property

Private Function IsDim(ByRef vArray As Variant) As Boolean

    On Error Resume Next
    IsDim = IsNumeric(UBound(vArray))
    On Error GoTo 0

End Function

Public Property Get List(ByVal vlngIndex As Long) As String

    List = mstrItemArray(vlngIndex)

End Property

Public Property Get Sorted() As Boolean

    Sorted = mblnSorted

End Property

Public Property Let Sorted(ByVal vblnNewValue As Boolean)

    mblnSorted = vblnNewValue

End Property

Public Function FindItem(ByVal vstrItem As String, _
                         Optional ByVal vblnInstr As Boolean = False) As Long

  Dim lngI As Long

    For lngI = 0 To UBound(mstrItemArray)
        If Not vblnInstr Then
            '/* Look for exact match
            If LCase$(mstrItemArray(lngI)) = LCase$(vstrItem) Then
                FindItem = lngI
                Exit Function
            End If
         Else
            '/* Look for Like match
            If InStr(LCase$(mstrItemArray(lngI)), LCase$(vstrItem)) <> 0 Then
                If Mid$(LCase$(mstrItemArray(lngI)), InStr(LCase$(mstrItemArray(lngI)), LCase$(vstrItem)), Len(vstrItem)) = LCase$(vstrItem) Then
                    FindItem = lngI
                    Exit Function
                End If
            End If
        End If
    Next lngI

End Function
