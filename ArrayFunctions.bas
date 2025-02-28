' Transfers the contents of an array to a specified range in the worksheet.
Sub ArrayToRange(arr As Variant, rng As Range)
    rng.Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
End Sub

' Transfers the contents of a specified range in the worksheet to an array.
Function RangeToArray(rng As Range) As Variant
    RangeToArray = rng.Value
End Function

' Resizes an array to the specified number of rows and columns.
Function ArrayResize(arr As Variant, newRows As Long, newCols As Long) As Variant
    Dim newArr() As Variant
    ReDim newArr(1 To newRows, 1 To newCols)
    Dim i As Long, j As Long
    For i = 1 To UBound(arr, 1)
        For j = 1 To UBound(arr, 2)
            If i <= newRows And j <= newCols Then
                newArr(i, j) = arr(i, j)
            End If
        Next j
    Next i
    ArrayResize = newArr
End Function

' Sorts a 2D array based on the values in a specified column.
Function ArraySort(arr As Variant, colIndex As Long, ascending As Boolean) As Variant
    Dim temp As Variant
    Dim i As Long, j As Long
    For i = 1 To UBound(arr, 1) - 1
        For j = i + 1 To UBound(arr, 1)
            If (ascending And arr(i, colIndex) > arr(j, colIndex)) Or (Not ascending And arr(i, colIndex) < arr(j, colIndex)) Then
                temp = arr(i, colIndex)
                arr(i, colIndex) = arr(j, colIndex)
                arr(j, colIndex) = temp
            End If
        Next j
    Next i
    ArraySort = arr
End Function

' Filters a 2D array based on a specified column and criteria.
Function ArrayFilter(arr As Variant, colIndex As Long, criteria As Variant) As Variant
    Dim filteredArr() As Variant
    ReDim filteredArr(1 To UBound(arr, 1), 1 To UBound(arr, 2))
    Dim i As Long, j As Long, k As Long
    k = 1
    For i = 1 To UBound(arr, 1)
        If arr(i, colIndex) = criteria Then
            For j = 1 To UBound(arr, 2)
                filteredArr(k, j) = arr(i, j)
            Next j
            k = k + 1
        End If
    Next i
    ReDim Preserve filteredArr(1 To k - 1, 1 To UBound(arr, 2))
    ArrayFilter = filteredArr
End Function

' Concatenates two arrays.
Function ArrayConcat(arr1 As Variant, arr2 As Variant) As Variant
    Dim newArr() As Variant
    ReDim newArr(1 To UBound(arr1, 1) + UBound(arr2, 1), 1 To UBound(arr1, 2))
    Dim i As Long, j As Long
    For i = 1 To UBound(arr1, 1)
        For j = 1 To UBound(arr1, 2)
            newArr(i, j) = arr1(i, j)
        Next j
    Next i
    For i = 1 To UBound(arr2, 1)
        For j = 1 To UBound(arr2, 2)
            newArr(i + UBound(arr1, 1), j) = arr2(i, j)
        Next j
    Next i
    ArrayConcat = newArr
End Function

' Returns an array containing only the unique values from the input array.
Function ArrayUnique(arr As Variant) As Variant
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To UBound(arr)
        If Not dict.exists(arr(i)) Then
            dict.Add arr(i), Nothing
        End If
    Next i
    ArrayUnique = dict.keys
End Function

' Transposes a 2D array (rows become columns and columns become rows).
Function ArrayTranspose(arr As Variant) As Variant
    Dim newArr() As Variant
    ReDim newArr(1 To UBound(arr, 2), 1 To UBound(arr, 1))
    Dim i As Long, j As Long
    For i = 1 To UBound(arr, 1)
        For j = 1 To UBound(arr, 2)
            newArr(j, i) = arr(i, j)
        Next j
    Next i
    ArrayTranspose = newArr
End Function
