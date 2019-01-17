---
title: Range.Item property (Excel)
keywords: vbaxl10.chm144151
f1_keywords:
- vbaxl10.chm144151
ms.prod: excel
api_name:
- Excel.Range.Item
ms.assetid: f7d40273-5069-8a9d-14ee-19df225f864c
ms.date: 06/08/2017
localization_priority: Priority
---


# Range.Item property (Excel)

Returns a **Range** object that represents a range at an offset to the specified range.


## Syntax

_expression_.**Item**(_RowIndex_, _ColumnIndex_)

_expression_ A variable that represents a [Range](Excel.Range(object).md) object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _RowIndex_|Required| **Variant**|The index number of the cell that you want to access, in order from left to right, and then down.<br/><br/>`Range.Item(1)` returns the upper-left cell in the range.<br/><br/>`Range.Item(2)` returns the cell immediately to the right of the upper-left cell. |
| _ColumnIndex_|Optional| **Variant**|A number or string that indicates the column number of the cell you want to access, starting with either 1 or "A" for the first column in the range.|

## Remarks

Syntax 1 uses a row number and a column number or letter as index arguments. For more information about this syntax, see the **Range** object. 

The **RowIndex** and **ColumnIndex** arguments are relative offsets. In other words, specifying a **RowIndex** of 1 returns cells in the first row of the range, not the first row of the worksheet. For example, if the selection is cell C3, `Selection.Cells(2, 2)` returns cell D4 (you can use the **Item** property to index outside the original range).


## Example

This example fills the range A1:A10 on Sheet1, based on the contents of cell A1. The **[Cells](excel.range.cells.md)** property returns a **Range** object.

```vb
Worksheets("Sheet1").Cells.Item("A1:A10").FillDown.
```



