<div align="center">

## Grid to Excel


</div>

### Description

This SubRoutine will print the MSHFlexGrid Content to Excel as it is along with giving borders,colors,bold. Its quiet a small function but useful sometimes.
 
### More Info
 
The MSHflexgrid Name which holds the data

Reference has to be set to Excel

Transfers the Grid Data to Excel


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rachit K](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rachit-k.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VBA MS Excel
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rachit-k-grid-to-excel__1-47408/archive/master.zip)





### Source Code

```
Private Sub Grid2Excel(gridName As MSHFlexGrid)
'This is the function to print from the Grid to Excel
Dim exc As Excel.Application
Set exc = CreateObject("Excel.Application")
exc.Workbooks.Add
exc.Visible = True
With gridName
  For i = 0 To .Rows - 1
    For j = 1 To .Cols - 1
      exc.Cells(i + 1, j) = .TextMatrix(i, j)
      exc.Cells(i + 1, j).Borders.LineStyle = xlDouble
      exc.Cells(i + 1, j).Borders.Color = vbBlue
    Next j
  Next i
  exc.Range("A1:" & Chr(65 + j) & 1).Font.Bold = True
  exc.Columns("$A:" & "$" & Chr(65 + j)).AutoFit
End With
End Sub
```

