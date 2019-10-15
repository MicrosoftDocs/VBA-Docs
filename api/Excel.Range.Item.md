---
title: Range.Item property (Excel)
keywords: vbaxl10.chm144151
f1_keywords:
- vbaxl10.chm144151
ms.prod: excel
api_name:
- Excel.Range.Item
ms.assetid: f7d40273-5069-8a9d-14ee-19df225f864c
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.Item property (Excel)

Returns a **Range** object that represents a range at an offset to the specified range.


## Syntax

_expression_.**Item** (_RowIndex_, _ColumnIndex_)

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _RowIndex_|Required| **Variant**|If the second argument is provided, the relative row number of the cell to return.<br/><br/>If the second argument is not provided, the index of the subrange to return. |
| _ColumnIndex_|Optional| **Variant**|The relative column number of the cell to return.|

## Remarks

If _expression_ is not a range containing a collection of single cells, e.g. because is has been obtained via the `[Columns](Excel.Range.Columns.md)` member, providing the second argument is illegal and will result in an error 1004. 

The default member of **Range** forwards calls with parameters to the **Item** member. Thus, `someRange(1)` and `someRange(1,1)` are equivalent to `someRange.Item(1)` and `someRange.Item(1,1)`, respectively.

The _RowIndex_ and _ColumnIndex_ arguments are relative offsets. In other words, specifying a _RowIndex_ of 1 returns cells in the first row of the range, not the first row of the worksheet. For example, if the selection is cell C3, `Selection.Cells(2, 2)` returns cell D4 (you can use the **Item** property to index outside the original range).

The enumeration order when supplying one parameter agrees with that used when enumerating the range in a `For Each` loop. For ranges consisting of single cells, as returned by the `[Cells](Excel.Range.Cells.md)` member, this is left-to-right than top-to-bottom, as for two-dimentional arrays. For ranges consisting of row ranges, as returned by `[Rows](Excel.Range.Rows.md)`, it is top-to-bottom and for ranges consisting of column ranges, as returned by `[Columns](Excel.Range.Columns.md)`, it is left-to-right.

If _expression_ is a range consisting of column ranges, the parameter _RowIndex_ refers to the relative column index and not to the relative row index.


## Example

The following example shows which cell is returned if both parameters are provided. 

```vba 
Public Sub PrintAdresses()
  Dim exampleRange As Excel.Range
  Set exampleRange = ThisWorkbook.Worksheets("ExampleSheet").Range("B2:D4")
  
  Debug.Print exampleRange.Item(1,1).Address  'Prints "$B$2"
  Debug.Print exampleRange.Item(2,4).Address  'Prints "$E$3"
End Sub
```

The following example shows for different types of ranges which subranges are returned if only one parameter is provided.

```vba 
Public Sub PrintAdresses()
  Dim exampleRange As Excel.Range
  Set exampleRange = ThisWorkbook.Worksheets("ExampleSheet").Range("B2:D4")
  
  Debug.Print exampleRange.Cells.Item(1).Address  'Prints "$B$2"
  Debug.Print exampleRange.Cells.Item(2).Address  'Prints "$C$2"
  Debug.Print exampleRange.Cells.Item(4).Address  'Prints "$B$3"
  
  Debug.Print exampleRange.Rows.Item(1).Address  'Prints "$B$2:$D$2"
  Debug.Print exampleRange.Rows.Item(2).Address  'Prints "$B$3:$D$3"
  
  Debug.Print exampleRange.Columns.Item(1).Address  'Prints "$B$2:$B$4"
  Debug.Print exampleRange.Columns.Item(2).Address  'Prints "$C$1:$C$4"
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
