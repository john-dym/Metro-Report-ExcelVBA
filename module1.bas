Option Explicit
Sub A01_Metro_Report()

'Find the total sheet size and assign active workbook/worksheet to variables
Dim maxCol As Long, maxRow As Long
Dim wb As Workbook, ws As Worksheet

Set wb = ActiveWorkbook
Set ws = ActiveSheet
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
Dim arrIndexes(1 To 5) As Long

For i = 1 To maxCol
    If Cells(1, i).Value = partNoLabel Then
        cPartNo = i
        arrIndexes(1) = i
    End If
    If Cells(1, i).Value = partNameLabel Then
        cPartName = i
        arrIndexes(2) = i
    End If
    If Cells(1, i).Value = partEoLabel Then
        cPartEo = i
        arrIndexes(3) = i
    End If
    If Cells(1, i).Value = partLotLabel Then
        cPartLot = i
        arrIndexes(4) = i
    End If
    If Cells(1, i).Value = partLocLabel Then
        cPartLoc = i
        arrIndexes(5) = i
    End If
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

'Grabs door labels and location numbers from ThisWorkbook
Dim maxDoors As Long
'ThisWorkbook.Windows(1).Visible = True
With ThisWorkbook.Worksheets(1)
    .Activate
    Range("A65530").Select
    Selection.End(xlUp).Select
    maxDoors = ActiveCell.Row - 1
End With

Dim doors() As String
ReDim doors(maxDoors)
    For i = 1 To maxDoors
        doors(i) = Cells(i + 1, 1).Value
    Next i

Dim locs() As String
ReDim locs(maxDoors)
    For i = 1 To maxDoors
        locs(i) = Cells(i + 1, 2).Value
    Next i

'ThisWorkbook.Windows(1).Visible = False
wb.Activate
ws.Activate

'Determine origTable array size
'Dim arraySize As Long
'arraySize = 0
 '   For i = 1 To maxDoors
  '      For j = 1 To maxRow
   '         If Cells(j, cPartLoc).Value = locs(i) Then
    '            arraySize = arraySize + 1
     '       End If
      '  Next j
    'Next i

'Parse through list and write matching locations' columns to array
Dim origTable() As String
ReDim origTable(5, 1)

For i = 1 To maxDoors
    For j = 1 To maxRow
        If Cells(j, cPartLoc).Value = locs(i) Then
            For k = 1 To 5
                origTable(k, UBound(origTable, 2)) = Cells(j, arrIndexes(k)).Value
            Next k
            ReDim Preserve origTable(5, UBound(origTable, 2) + 1)
        End If
    Next j
Next i

End Sub
