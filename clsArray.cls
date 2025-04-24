VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'@IgnoreModule ArgumentWithIncompatibleObjectType, ModuleWithoutFolder
Private anArray() As Variant
Private arrType As Integer

Public Property Get Members() As Variant() 'Get the members of the array
    Members = anArray
End Property

Public Property Let Members(ByRef arr() As Variant) 'Add full array to Class Object
    anArray = arr
End Property

Property Let Member(index As Long, inValue As Variant) 'Let the value of the specified member of the array = the specified value
    anArray(index) = inValue
End Property

Property Get Member(index As Long) As Variant 'Get a specified member of 1D array
    Member = anArray(index)
End Property

Property Let Member2D(index1 As Long, index2 As Long, inValue As Variant) 'Let the value of the specified member of the array = the specified value
    anArray(index1, index2) = inValue
End Property

Property Get Member2D(index1 As Long, index2 As Long) As Variant 'Get a specified member of 2D array
    Member2D = anArray(index1, index2)
End Property



Public Sub AddToMember(index As Long, value As Long) 'Adds value to array index value
    Dim X As Variant
    X = anArray(index)
    If X = Empty Then
        X = 0
    End If
    anArray(index) = X + value
End Sub

Public Sub AddToMember2D(index1 As Long, index2 As Long, value As Variant) 'Adds value to array index value
    Dim X As Variant
    X = anArray(index1, index2)
    If X Is Empty Then
        X = 0
    End If
    anArray(index1, index2) = X + value
End Sub

Property Let ArrayType(dimensions As Integer) 'Note the ArrayType by the number of dimensions 1 or 2
    arrType = dimensions
End Property

Property Get ArrayType() As Integer 'Get the current ArrayType dimensions
    ArrayType = arrType
End Property

Public Sub SetArr(index As Long) 'Set the bounds of the Array
    ReDim anArray(1 To index)
End Sub

Public Sub SetArr2D(index1 As Long, index2 As Long) 'Set the bounds of the Array
    ReDim anArray(1 To index1, 1 To index2)
End Sub

Public Sub Init(ByRef iArray() As Variant) 'Add full array to Class Object
    anArray = iArray
End Sub


Property Get MinValue() As Long 'Property to get Min of 1D Array
    MinValue = Application.Min(anArray)
End Property


Property Get MaxValue() As Long 'Property to get Max of 1D Array
    MaxValue = Application.Max(anArray)
End Property

'— flatten a 2D 1×N or N×1 array into a 1D (1 To N) array —
Private Function Flatten2Dto1D(arr2D As Variant) As Variant
    Dim r As Long, c As Long
    r = UBound(arr2D, 1)
    On Error Resume Next
        c = UBound(arr2D, 2)
    On Error GoTo 0

    Dim out() As Variant
    If r = 1 Then
        ' row vector 1×c
        ReDim out(1 To c)
        Dim j As Long
        For j = 1 To c
            out(j) = arr2D(1, j)
        Next j
    Else
        ' column vector r×1
        ReDim out(1 To r)
        Dim i As Long
        For i = 1 To r
            out(i) = arr2D(i, 1)
        Next i
    End If

    Flatten2Dto1D = out
End Function


Property Get ArrColumn(index As Long) As Variant()
    Dim col2D As Variant
    col2D = Application.Index(anArray, 0, index)    ' returns an r×1 array
    ArrColumn = Flatten2Dto1D(col2D)
End Property

Property Get ArrRow(index As Long) As Variant()
    Dim row2D As Variant
    row2D = Application.index(anArray, index, 0)    ' returns a 1×c array
    ArrRow = row2D
End Property


Property Get ArrColumns(colIdx() As Variant) As Variant()
    Dim r As Long: r = UBound(anArray, 1)
    Dim c As Long: c = UBound(colIdx)
    Dim out() As Variant
    ReDim out(1 To r, 1 To c+1)

    Dim i As Long, j As Long
    For j = 1 To c + 1
        For i = 1 To r
            out(i, j) = anArray(i, colIdx(j-1))
        Next i
    Next j

    ArrColumns = out
End Property


Property Get ArrSize(Optional dimension As Long) As Long
    dimension = IIf(dimension = Empty, 1, dimension)
    On Error GoTo Err0
    ArrSize = UBound(anArray, dimension)
    Exit Property
Err0:
    ArrSize = 0
End Property

Public Sub EraseArr() 'Sub to Erase array
    Erase anArray
End Sub

Public Function RemoveTopRow() As Variant()
    If UBound(anArray, 1) >= 1 Then
        ' Get the number of rows and columns in the current array
        Dim NumRows As Long
        Dim numCols As Long
        NumRows = UBound(anArray, 1) - LBound(anArray, 1) + 1
        numCols = UBound(anArray, 2) - LBound(anArray, 2) + 1

        ' Check if there is more than one row
        If NumRows > 1 Then
            ' Create a new array to store the result with one less row
            Dim resultArray() As Variant
            ReDim resultArray(1 To NumRows - 1, 1 To numCols)

            ' Copy the remaining rows (excluding the top row) to the result array
            Dim i As Long
            For i = 2 To NumRows
                Dim j As Long
                For j = 1 To numCols
                    resultArray(i - 1, j) = anArray(i, j)
                Next j
            Next i

            ' Update the original array with the result
            anArray = resultArray
        Else
            ' If there's only one row, clear the array
            Erase anArray
        End If
    End If

    ' Return the result array (or an empty array if no rows remain)
    RemoveTopRow = anArray
End Function


Public Function RemoveFirstItem() As Variant()
    If UBound(anArray, 1) >= LBound(anArray, 1) Then
        ' Get the number of items in the current one-dimensional array
        Dim numItems As Long
        numItems = UBound(anArray, 1) - LBound(anArray, 1) + 1

        ' Check if there is more than one item
        If numItems > 1 Then
            ' Create a new array to store the result with one less item
            Dim resultArray() As Variant
            ReDim resultArray(1 To numItems - 1)

            ' Copy the remaining items (excluding the first item) to the result array
            Dim i As Long
            For i = 2 To numItems
                resultArray(i - 1) = anArray(i)
            Next i

            ' Update the original array with the result
            anArray = resultArray
        Else
            ' If there's only one item, clear the array
            Erase anArray
        End If
    Else
        ' Handle the case where it's an empty array
        ' In this case, just clear the array
        Erase anArray
    End If

    ' Return the result array (or an empty array if no items remain)
    RemoveFirstItem = anArray
End Function

Public Function RemoveColumns(ParamArray columnIndices() As Variant) As Variant()
    If UBound(anArray, 2) >= LBound(anArray, 2) Then
        ' Get the number of rows and columns in the current array
        Dim NumRows As Long
        Dim numCols As Long
        NumRows = UBound(anArray, 1) - LBound(anArray, 1) + 1
        numCols = UBound(anArray, 2) - LBound(anArray, 2) + 1

        If numCols >= 1 Then
            ' Create an array to store the column indices to be removed
            Dim columnsToRemove() As Boolean
            ReDim columnsToRemove(1 To numCols)

            ' Initialize columnsToRemove to False (not to remove) for all columns
            Dim i As Long
            For i = 1 To numCols
                columnsToRemove(i) = False
            Next i

            ' Mark the specified column indices to be removed
            Dim columnIndex As Variant
            For Each columnIndex In columnIndices
                If IsNumeric(columnIndex) Then
                    Dim colIndex As Long
                    colIndex = CLng(columnIndex)

                    If colIndex >= 1 And colIndex <= numCols Then
                        columnsToRemove(colIndex) = True
                    End If
                End If
            Next columnIndex

            ' Calculate the number of columns to keep
            Dim numColsToKeep As Long
            numColsToKeep = Application.WorksheetFunction.CountIf(columnsToRemove, False)

            If numColsToKeep > 0 Then
                ' Create a new array to store the result with the specified columns removed
                Dim resultArray() As Variant
                ReDim resultArray(1 To NumRows, 1 To numColsToKeep)

                ' Copy the remaining columns (excluding the specified columns) to the result array
                Dim newRow As Long
                For newRow = 1 To NumRows
                    Dim newCol As Long
                    Dim targetCol As Long
                    targetCol = 1
                    For newCol = 1 To numCols
                        If Not columnsToRemove(newCol) Then
                            resultArray(newRow, targetCol) = anArray(newRow, newCol)
                            targetCol = targetCol + 1
                        End If
                    Next newCol
                Next newRow

                ' Update the original array with the result
                anArray = resultArray
            Else
                ' If all specified columns are removed, clear the array
                Erase anArray
            End If
        End If
    End If

    ' Return the result array (or an empty array if no columns remain)
    RemoveColumns = anArray
End Function

Public Function QuickFilter(column_index As Long, filter_value As Variant, Optional eq As String) As Variant()
    Dim r As Long
    Dim Count As Long
    Dim Indices As Variant
    On Error GoTo Err
    ReDim Indices(LBound(anArray, 1) To UBound(anArray, 1))
    Count = LBound(anArray, 1) - 1
    If eq = "N" Then
        For r = LBound(anArray, 1) To UBound(anArray, 1)
            If anArray(r, column_index) <> filter_value Then
                Count = Count + 1
                Indices(Count) = r
            End If
        Next
    Else
        For r = LBound(anArray, 1) To UBound(anArray, 1)
            If anArray(r, column_index) = filter_value Then
                Count = Count + 1
                Indices(Count) = r
            End If
        Next
    End If
    
    Dim Results As Variant
    Dim c As Long
    Dim index As Long
    If Count >= LBound(anArray, 1) Then
'        we want results to hold the entire row
        ReDim Results(LBound(anArray, 1) To Count, LBound(anArray, 2) To UBound(anArray, 2))
        For r = LBound(anArray, 1) To Count
            index = Indices(r)
            For c = LBound(anArray, 2) To UBound(anArray, 2)
                Results(r, c) = anArray(index, c)
            Next
        Next
        QuickFilter = Results
    Else
        QuickFilter = Array()
    End If
Exit Function

Err:
    QuickFilter = anArray

End Function

Public Function ArraySearch(lookup As Variant, lookup_col As Long, return_col As Long, Optional blank_value As String, Optional match_mode As String = "Exact") As Variant
    Dim r As Long
    Dim Count As Long
    Dim Indices As Variant
    On Error GoTo Err
    ReDim Indices(LBound(anArray, 1) To UBound(anArray, 1))

    If match_mode = "Exact" Then
        For r = LBound(anArray, 1) To UBound(anArray, 1)
            If anArray(r, lookup_col) = lookup Then
                ArraySearch = anArray(r, return_col)
                Exit Function
            Else
                ArraySearch = ArraySearch(lookup, lookup_col, return_col, "", "N")
                Exit Function
            End If
        Next
    Else
        For r = LBound(anArray, 1) To UBound(anArray, 1)
            If InStr(1, anArray(r, lookup_col), lookup, 0) > 0 Then
                ArraySearch = anArray(r, return_col)
                Exit Function
            Else
                ArraySearch = blank_value
            End If
        Next
    End If
Exit Function

Err:
    ArraySearch = blank_value

End Function

Public Function Array_Index(lookup As Variant, lookup_col As Long, Optional blank_value As String, Optional match_mode As String = "Exact") As Variant
    Dim r As Long
    Dim Count As Long
    Dim Indices As Variant
    On Error GoTo Err
    ReDim Indices(LBound(anArray, 1) To UBound(anArray, 1))

    If match_mode = "Exact" Then
        For r = LBound(anArray, 1) To UBound(anArray, 1)
            If anArray(r, lookup_col) = lookup Then
                Array_Index = r
                Exit Function
            Else
                Array_Index = Array_Index(lookup, lookup_col, "", "N")
                Exit Function
            End If
        Next
    Else
        For r = LBound(anArray, 1) To UBound(anArray, 1)
            If InStr(1, anArray(r, lookup_col), lookup, 0) > 0 Then
                Array_Index = r
                Exit Function
            Else
                Array_Index = blank_value
            End If
        Next
    End If
Exit Function

Err:
    Array_Index = blank_value
    
End Function

Function JoinArrays(newRows As Variant) As Variant()
    ' Get the dimensions of the original array
    Dim originalRows As Long
    Dim originalColumns As Long
    On Error GoTo Empty_anArray
    originalRows = UBound(anArray, 1)
    originalColumns = UBound(anArray, 2)


    ' Get the dimensions of the new rows array
    Dim newRowsCount As Long
    Dim newRowsColumns As Long

    On Error GoTo 0
    newRowsCount = UBound(newRows, 1)
    newRowsColumns = UBound(newRows, 2)

    ' Check if the column count matches between the original array and the new rows
    If originalColumns <> newRowsColumns Then
        MsgBox "Column count mismatch between original array and new rows."
        Exit Function
    End If

    ' Create a new array with the combined rows
    Dim combinedArray() As Variant
    ReDim combinedArray(1 To originalRows + newRowsCount, 1 To originalColumns)

    ' Copy the original array into the new combined array
    Dim i As Long, j As Long
    For i = 1 To originalRows
        For j = 1 To originalColumns
            combinedArray(i, j) = anArray(i, j)
        Next j
    Next i

    ' Append the new rows to the combined array
    Dim newRowIndex As Long
    For i = 1 To newRowsCount
        For j = 1 To newRowsColumns
            combinedArray(originalRows + i, j) = newRows(i, j)
        Next j
    Next i

    ' Return the combined array
    JoinArrays = combinedArray
    Exit Function

Empty_anArray:
    JoinArrays = newRows
    Exit Function

Empty_newRow:
    Debug.Print "Empty New Row"
    Exit Function

End Function

Sub JoinColumn(newColumns As Variant)

    Dim currentColumnNum    as Long: currentColumnNum = ArrSize(2)
    Dim newColumnsNum       as Long: newColumnsNum = UBound(newColumns, 2)
    Dim totalNewColumns     as Long: totalNewColumns = currentColumnNum + newColumnsNum
    Dim relativeColNum      as Long

    ReDim Preserve anArray(1 To ArrSize(1), 1 To totalNewColumns)

    Dim r as Long
    Dim c as Long

    For r = 1 to ArrSize(1)
        For c = 1 to newColumnsNum
            relativeColNum = currentColumnNum + c
            Member2D(r, relativeColNum) = newColumns(r, c)
        Next
    Next

End Sub


' Method to set a worksheet range with the values in anArray
Public Sub SetRange(ws As Worksheet, topLeftCell As String, Optional clear_cells As Boolean = False, Optional additonal_rows As Long = 0)
    Dim rowCount As Long
    Dim columnCount As Long
    
    ' Get the number of rows and columns in the internal array
    rowCount = UBound(anArray, 1) - LBound(anArray, 1) + 1
    columnCount = UBound(anArray, 2) - LBound(anArray, 2) + 1
    
    ' Determine the range to be set
    Dim targetRange As Range
    Set targetRange = ws.Range(topLeftCell).Resize(rowCount, columnCount)

    If clear_cells Then
        Dim oldRange As Range: Set oldRange = ws.Range(topLeftCell).Resize(rowCount + additonal_rows, columnCount)
        oldRange.ClearContents
    End If
    
    ' Set the range to the values in anArray
    targetRange.value = anArray
End Sub


Public Sub SetListObjectData(listObject As listObject)
    Dim rowCount As Long
    Dim columnCount As Long
    
    ' Get the number of rows and columns in the internal array
    rowCount = UBound(anArray, 1) - LBound(anArray, 1) + 1
    columnCount = UBound(anArray, 2) - LBound(anArray, 2) + 1
    
    ' Check if the column count matches the ListObject's number of columns
    If columnCount <> listObject.ListColumns.Count Then
        MsgBox "Column count mismatch between array and ListObject."
        Exit Sub
    End If
    
    ' Clear the existing data in the ListObject
    If Not listObject.DataBodyRange Is Nothing Then
        listObject.DataBodyRange.Delete
    End If
    
    ' Resize the ListObject to match the row count of the array
    Dim targetRange As Range
    Set targetRange = listObject.HeaderRowRange.Offset(1, 0).Resize(rowCount, columnCount)
    
    ' Set the range with the internal array values
    targetRange.value = anArray
End Sub

Public Function GetUniqueValues(columnIndex As Long) As Variant()
    Dim colArray As Variant
    Dim uniqueDict As Object
    Dim i As Long
    Dim result() As Variant
    Dim key As Variant
    
    ' Extract the entire column using ArrColumn property
    colArray = Me.ArrColumn(columnIndex)
    
    ' Create dictionary for uniqueness check
    Set uniqueDict = CreateObject("Scripting.Dictionary")
    
    ' Loop through the column values and store unique ones
    For i = LBound(colArray) To UBound(colArray)
        If Not uniqueDict.exists(colArray(i)) Then
            uniqueDict.Add colArray(i), colArray(i)
        End If
    Next i
    
    ' Convert dictionary keys to array
    result = uniqueDict.keys
    
    ' Return unique values
    GetUniqueValues = result
End Function


Public Function Pivot_Array(group_cols As Variant, column_values As Variant, column_col As Long, sum_col As Long) As Variant()
    Dim uniqueGroups As Object
    Dim row As Long, col As Long
    Dim groupKey As String
    Dim resultArray() As Variant
    Dim resultRow As Long
    Dim key As Variant
    Dim i As Integer
    Dim numGroupCols As Integer
    Dim colIndex As Integer
    Dim sumValue As Variant
    Dim columnValue As Variant

    ' Create dictionary to store grouped values
    Set uniqueGroups = CreateObject("Scripting.Dictionary")
    
    ' Determine number of grouping columns
    numGroupCols = UBound(group_cols) - LBound(group_cols) + 1
    
    ' Loop through the array and aggregate data
    For row = LBound(anArray, 1) To UBound(anArray, 1)
        ' Create a unique key for grouping columns
        groupKey = ""
        For col = LBound(group_cols) To UBound(group_cols)
            groupKey = groupKey & anArray(row, group_cols(col)) & "_"
        Next col
        groupKey = Left(groupKey, Len(groupKey) - 1) ' Remove trailing "_"
        
        ' Get column value and sum value
        columnValue = anArray(row, column_col)
        sumValue = anArray(row, sum_col)
        
        ' Initialize if key does not exist
        If Not uniqueGroups.exists(groupKey) Then
            ReDim tempArray(1 To numGroupCols + UBound(column_values) - LBound(column_values) + 1)
            ' Store grouping column values
            For i = 1 To numGroupCols
                tempArray(i) = anArray(row, group_cols(LBound(group_cols) + i - 1))
            Next i
            ' Initialize column values to 0
            For i = numGroupCols + 1 To numGroupCols + UBound(column_values) - LBound(column_values) + 1
                tempArray(i) = 0
            Next i
            uniqueGroups.Add groupKey, tempArray
        End If
        
        ' Retrieve the array, modify it, and store it back
        tempArray = uniqueGroups(groupKey)
        ' Find the correct column index based on provided column values
        For i = LBound(column_values) To UBound(column_values)
            If columnValue = column_values(i) Then
                colIndex = numGroupCols + i - LBound(column_values) + 1
                tempArray(colIndex) = tempArray(colIndex) + sumValue
                Exit For
            End If
        Next i
        
        ' Store the updated array back into the dictionary
        uniqueGroups(groupKey) = tempArray
        
    Next row
    
    ' Create result array without headers
    resultRow = 1
    ReDim resultArray(1 To uniqueGroups.Count, 1 To numGroupCols + UBound(column_values) - LBound(column_values) + 1)

    ' Populate result array
    For Each key In uniqueGroups.keys
        For i = 1 To numGroupCols + UBound(column_values) - LBound(column_values) + 1
            resultArray(resultRow, i) = uniqueGroups(key)(i)
        Next i
        resultRow = resultRow + 1
    Next key

    ' Return the pivoted result
    Pivot_Array = resultArray
End Function

