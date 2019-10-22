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

If _expression_ is not a range containing a collection of single cells, e.g. because is has been obtained via the **[Columns](Excel.Range.Columns.md)** member, providing the second argument is illegal and will result in an error 1004. 

The default member of **Range** forwards calls with parameters to the **Item** member. Thus, `someRange(1)` and `someRange(1,1)` are equivalent to `someRange.Item(1)` and `someRange.Item(1,1)`, respectively.

The _RowIndex_ and _ColumnIndex_ arguments are relative 1-based offsets to the top-left cell of the first area of the range as returned by the **[Areas](Excel.Range.Areas.md)** member, i.e. for the range `Union(someSheet.Range("Z4:AA6"), someSheet.Range("A1:C3"))`, `Item(1,1)` will return the range with address $Z$4.

The _ColumnIndex_ can be provided either as a numeric index or as a column address string as in A1-notation, i.e. `"A"` refers to the numeric index `1` and `"AA"` to `27`. 

It is possible to reference cells outside the original range using the **Item** property by providing appropriate arguments, e.g. `Item(3,3)` will return the cell at `"D4"` for the range `someSheet.Range("B2:C3")`. 

The range returned when providing only one parameter depends on the nature of the range: 

  - For ranges consisting of single cells, as returned by the **[Cells](Excel.Range.Cells.md)** and **[Range](Excel.Worksheet.Range.md)** members, **Item** returns single cells. The parameter _RowIndex_ refers to the index when enumerating the first area of the range left-to-right than top-to-bottom, as for two-dimentional arrays. If _RowIndex_ is larger than the number of cells in the first area of the range, the enumeration as if the area was extended downwards. 

  - For ranges consisting of row ranges, as returned by **[Rows](Excel.Range.Rows.md)**, **Item** returns row ranges. The parameter _RowIndex_ is the 1-based offset from the first row of the range in the direction top-to-bottom. 
  
  - For ranges consisting of column ranges, as returned by **[Columns](Excel.Range.Columns.md)**, **Item** returns column ranges. The parameter _RowIndex_ is the 1-based offset from the first column in the direction left-to-right. In this situation, the index may alternatively be provided as a column address string.


## Example

The following example shows which cell is returned if both parameters are provided. 

```vb
Public Sub PrintAdresses()
   Dim exampleRange As Excel.Range
   With ThisWorkbook.Worksheets("ExampleSheet")
      Set exampleRange = Application.Union(.Range("B2:D4"), .Range("A1"), .Range("Z1:AA20"))
   End With
  
   Debug.Print exampleRange.Item(1,1).Address      'Prints "$B$2"
   Debug.Print exampleRange.Item(2,4).Address      'Prints "$E$3"
   Debug.Print exampleRange.Item(20,40).Address    'Prints "$AO$21"
   Debug.Print exampleRange.Item(2,"D").Address    'Prints "$E$3"
   Debug.Print exampleRange.Item(20,"AN").Address  'Prints "$E$3"
End Sub
```

The following example shows for different types of ranges which subranges are returned if only one parameter is provided.

```vb
Public Sub PrintAdresses()
   Dim exampleRange As Excel.Range
   With ThisWorkbook.Worksheets("ExampleSheet")
      Set exampleRange = Application.Union(.Range("B2:D4"), .Range("A1"), .Range("Z1:AA20"))
   End With

   Debug.Print exampleRange.Cells.Item(1).Address      'Prints "$B$2"
   Debug.Print exampleRange.Cells.Item(2).Address      'Prints "$C$2"
   Debug.Print exampleRange.Cells.Item(4).Address      'Prints "$B$3"
   Debug.Print exampleRange.Cells.Item(10).Address     'Prints "$B$5"
  
   Debug.Print exampleRange.Rows.Item(1).Address       'Prints "$B$2:$D$2"
   Debug.Print exampleRange.Rows.Item(10).Address      'Prints "$B$11:$D$11"
  
   Debug.Print exampleRange.Columns.Item(1).Address    'Prints "$B$2:$B$4"
   Debug.Print exampleRange.Columns.Item(10).Address   'Prints "$K$1:$K$4"
   Debug.Print exampleRange.Columns.Item("A").Address  'Prints "$B$1:$B$4"
   Debug.Print exampleRange.Columns.Item("J").Address  'Prints "$K$1:$K$4"
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
