Sub A01_Metro_Report()

'Find the total sheet size
Dim maxCol As Long, maxRow As Long

    Range("AZ1").Select
    Selection.End(xlToLeft).Select
    maxCol = ActiveCell.Column
    Range("A65530").Select
    Selection.End(xlUp).Select
    maxRow = ActiveCell.Row

'Column headers on the first row
Dim partNoLabel As String, partNameLabel As String, partLocLabel As String, partEoLabel As String, partLotLabel As String
Dim i As Integer
partNoLabel = "Part No"
partNameLabel = "Part Name"
partLocLabel = "Loc. No"
partEoLabel = "EO No"
partLotLabel = "LOT No"

'Find column cells that match the labels for cell references
Dim cPartNo As Long
Dim cPartName As Long
Dim cPartLoc As Long
Dim cPartEo As Long
Dim cPartLot As Long

For i = 1 To maxCol
    If Cells(1, i).Value = partNoLabel Then cPartNo = i
    If Cells(1, i).Value = partNameLabel Then cPartName = i
    If Cells(1, i).Value = partLocLabel Then cPartLoc = i
    If Cells(1, i).Value = partEoLabel Then cPartEo = i
    If Cells(1, i).Value = partLotLabel Then cPartLot = i
Next i

MsgBox "Debug"


End Sub
