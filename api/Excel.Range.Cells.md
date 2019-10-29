---
title: Range.Cells property (Excel)
keywords: vbaxl10.chm144091
f1_keywords:
- vbaxl10.chm144091
ms.prod: excel
api_name:
- Excel.Range.Cells
ms.assetid: 32a6ecc7-2366-2cec-1feb-0966241a435d
ms.date: 08/14/2019
localization_priority: Priority
---


# Range.Cells property (Excel)

Returns a **Range** object that represents the cells in the specified range.

[!include[Add-ins note](~/includes/addinsnote.md)]


## Syntax

_expression_.**Cells**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

The return value is a **Range** consisting of single cells, which allows to use the version of the **[Item](Excel.Range.Item.md)** with two parameters and lets `For Each` loops iterate over single cells.

Because the default member of **Range** forwards calls with parameters to the **[Item](Excel.Range.Item.md)** property, you can specify the row and column index immediately after the **Cells** keyword instead of an explicit call to **[Item](Excel.Range.Item.md)**.

Using **Cells** without an object qualifier is equivalent to **[ActiveSheet.Cells](Excel.Worksheet.Cells.md)**.

## Example

This example sets the font style for cells B2:D6 on Sheet1 of the active workbook to italic.

```vb
With Worksheets("Sheet1").Range("B2:Z100") 
   .Range(.Cells(1, 1), .Cells(5, 3)).Font.Italic = True
End With
```

<br/>

This example scans a column of data named _myRange_. If a cell has the same value as the cell immediately preceding it, the example displays the address of the cell that contains the duplicate data.

```vb
Set r = Range("myRange") 
For n = 2 To r.Rows.Count 
    If r.Cells(n-1, 1) = r.Cells(n, 1) Then 
        MsgBox "Duplicate data in " & r.Cells(n, 1).Address 
    End If 
Next
```

<br/>

This example demonstrates how **Cells** changes the behavior of the **[Item](Excel.Range.Item.md)** member.  

```vb
Public Sub PrintRangeAdresses
   Dim columnsRange As Excel.Range
   Set columnsRange = ThisWorkBook.Worksheets("exampleSheet").Range("B2:Z100").Columns
   
   Debug.Print columnsRange.Item(2).Address         'Prints "$C$2:$C$100" 
   Debug.Print columnsRange.Cells.Item(2).Address   'Prints "$C$2" 
   Debug.Print columnsRange.Cells.Item(2,1).Address 'Prints "$B$3"   
End Sub
```

<br/>

This example demonstrates how **Cells** changes the enumeration behavior.

```vb
Public Sub PrintAllRangeAdresses
   Dim columnsRange As Excel.Range
   Set columnsRange = ThisWorkBook.Worksheets("exampleSheet").Range("B2:C3").Columns
   
   Dim columnRange As Excel.Range
   For Each columnRange In columnsRange
      Debug.Print columnRange.Address   'Prints "$B$2:$B$3", "$C$2:$C$3"
   Next
   
   Dim cell As Excel.Range
   For Each cell In columnsRange.Cells
      Debug.Print cell.Address          'Prints "$B$2", "$C$2", "$B$3", "$C$3"
   Next  
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
