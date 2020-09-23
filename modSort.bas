Attribute VB_Name = "Sort"
'------------------------------------------------------------
' Module    : modSort.bas
' Purpose   : Sort routines
' Author    : unknown
' Source    : Planet Source Code
'------------------------------------------------------------
'
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" _
                      Alias "RtlMoveMemory" ( _
                      lpDest As Any, _
                      lpSource As Any, _
                      ByVal cbCopy As Long)

'------------------------------------------------------------
' Purpose   : Sorts an array using a BubbleSort
' Arguments : varArray - array to sort
'------------------------------------------------------------
Public Sub BubbleSortMe(varArray() As String)

    Dim I As Long
    Dim j As Long
    Dim l_Count As Long

    l_Count = UBound(varArray)

    For I = 0 To l_Count
        For j = I + 1 To l_Count
            If LCase(varArray(I)) > LCase(varArray(j)) Then
                ' Here's the juice!
                SwapStrings varArray(I), varArray(j)
            End If
        Next
    Next

End Sub

'------------------------------------------------------------
' Purpose   : Sorts an array using a QuickSort
' Arguments : varArray - array to sort
'             l_First  - starting element
'             l_Last   - ending element
'------------------------------------------------------------
'
Public Sub QuickSortMe(varArray() As String, _
                       Optional l_First As Long = -1, _
                       Optional l_Last As Long = -1)

    Dim l_Low As Long
    Dim l_Middle As Long
    Dim l_High As Long
    Dim v_Test As Variant

    If l_First = -1 Then
        l_First = LBound(varArray)
    End If

    If l_Last = -1 Then
        l_Last = UBound(varArray)
    End If

    If l_First < l_Last Then
        l_Middle = (l_First + l_Last) / 2
        v_Test = LCase(varArray(l_Middle))
        l_Low = l_First
        l_High = l_Last

        Do
            While LCase(varArray(l_Low)) < v_Test
                l_Low = l_Low + 1
            Wend

            While LCase(varArray(l_High)) > v_Test
                l_High = l_High - 1
            Wend

            If (l_Low <= l_High) Then
                SwapStrings varArray(l_Low), varArray(l_High)
                l_Low = l_Low + 1
                l_High = l_High - 1
            End If

        Loop While (l_Low <= l_High)

        If l_First < l_High Then
            QuickSortMe varArray, l_First, l_High
        End If

        If l_Low < l_Last Then
            QuickSortMe varArray, l_Low, l_Last
        End If

    End If

End Sub

'------------------------------------------------------------
' Purpose   : Use 'CopyMemory' API to swap two array elements
' Arguments : pbString1 - 1st element
'             pbString2 - 2nd element
'------------------------------------------------------------
'
Private Sub SwapStrings(pbString1 As String, pbString2 As String)

    Dim l_Hold As Long

    CopyMemory l_Hold, ByVal VarPtr(pbString1), 4
    CopyMemory ByVal VarPtr(pbString1), ByVal VarPtr(pbString2), 4
    CopyMemory ByVal VarPtr(pbString2), l_Hold, 4

End Sub
