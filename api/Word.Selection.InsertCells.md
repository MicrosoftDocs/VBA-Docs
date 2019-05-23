---
title: Selection.InsertCells method (Word)
keywords: vbawd10.chm158662870
f1_keywords:
- vbawd10.chm158662870
ms.prod: word
api_name:
- Word.Selection.InsertCells
ms.assetid: 461085a3-ae98-8028-5ad2-d5e22038c6db
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.InsertCells method (Word)

Adds cells to an existing table.


## Syntax

_expression_. `InsertCells`( `_ShiftCells_` )

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ShiftCells_|Optional| **WdInsertCells**|Specifies how to insert the cells into the existing columns and rows of the table.|

## Remarks

The number of cells inserted is equal to the number of cells in the selection. You can also insert cells by using the  **[Add](Word.Cells.Add.md)** method of the **Cells** object.


## Example

This example inserts new cells to the left of the selected cells, and then it surrounds the selected cells with a red, single-line border.


```vb
If Selection.Cells.Count >= 1 Then 
 Selection.InsertCells ShiftCells:=wdInsertCellsShiftRight 
 For Each aBorder In Selection.Borders 
 aBorder.LineStyle = wdLineStyleSingle 
 aBorder.ColorIndex = wdRed 
 Next aBorder 
End If
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]