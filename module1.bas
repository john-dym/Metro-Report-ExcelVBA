Sub A01_Metro_Report()

'Find the total sheet size
Dim maxCol As Long, maxRow As Long

Range("AZ1").Select
Selection.End(xlToLeft).Select
maxCol = ActiveCell.Column
Range("A65530").Select
Selection.End(xlUp).Select
maxRow = ActiveCell.Row

'Column headers on the first row (Planning to have variables pulled directly from worksheet for easier user change)
Dim partNoLabel As String, partNameLabel As String, partLocLabel As String, partEoLabel As String, partLotLabel As String

partNoLabel = "Part No"
partNameLabel = "Part Name"
partLocLabel = "Loc. No"
partEoLabel = "EO No"
partLotLabel = "LOT No"

'Find column index that match the labels for cell references
Dim cPartNo As Long, cPartName As Long, cPartLoc As Long, cPartEo As Long, cPartLot As Long
Dim i As Integer, j As Integer, k As Integer

For i = 1 To maxCol
    If Cells(1, i).Value = partNoLabel Then cPartNo = i
    If Cells(1, i).Value = partNameLabel Then cPartName = i
    If Cells(1, i).Value = partEoLabel Then cPartEo = i
    If Cells(1, i).Value = partLotLabel Then cPartLot = i
    If Cells(1, i).Value = partLocLabel Then cPartLoc = i
Next i

'Sort rows by Loc, Part No, and Part Name
    ActiveSheet.Sort.SortFields.Clear
    ActiveSheet.Sort.SortFields.Add2 _
        Key:=Range(Cells(1, cPartLoc), Cells(maxRow, cPartLoc)), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveSheet.Sort.SortFields.Add2 _
        Key:=Range(Cells(1, cPartNo), Cells(maxRow, cPartNo)), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveSheet.Sort.SortFields.Add2 _
        Key:=Range(Cells(1, cPartName), Cells(maxRow, cPartName)), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveSheet.Sort
        .SetRange Range(Cells(1, 1), Cells(maxRow, maxCol))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

'Parse through list and write matching locations' columns to array
Dim doors(1 To 3) As String 'Planning to pull info from worksheet directly
Dim maxDoors As Long
Dim arrIndexes(1 To 5) As Long
Dim l As Integer
l = 1
doors(1) = "A7"
doors(2) = "A8"
doors(3) = "A9"
maxDoors = 3
Dim locs(1 To 3) As String
locs(1) = "MET015"
locs(2) = "MET021"
locs(3) = "MET031"

arrIndexes(1) = cPartNo
arrIndexes(2) = cPartName
arrIndexes(3) = cPartEo
arrIndexes(4) = cPartLot
arrIndexes(5) = cPartLoc

Dim origTable(1 To 2000, 1 To 5)  As String

For i = 1 To maxDoors
    For j = 1 To maxRow
        If Cells(j, cPartLoc).Value = locs(i) Then
            For k = 1 To 5
                origTable(l, k) = Cells(j, arrIndexes(k)).Value
                Next k
            l = l + 1
        End If
    Next j
Next i

End Sub
