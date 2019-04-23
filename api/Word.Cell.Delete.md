---
title: Cell.Delete method (Word)
keywords: vbawd10.chm156106952
f1_keywords:
- vbawd10.chm156106952
ms.prod: word
api_name:
- Word.Cell.Delete
ms.assetid: 01e6d989-e86c-9a3b-b0e3-d6eb1f2a7183
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell.Delete method (Word)

Deletes a table cell or cells and optionally controls how the remaining cells are shifted.


## Syntax

_expression_.**Delete**( `_ShiftCells_` )

_expression_ Required. A variable that represents a '[Cell](Word.Cell.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ShiftCells_|Optional| **Variant**|The direction in which the remaining cells are to be shifted. Can be any  **[WdDeleteCells](Word.WdDeleteCells.md)** constant. If omitted, cells to the right of the last deleted cell are shifted left.|

## Example

This example deletes the first cell in the first table of the active document.


```vb
Sub DeleteCells() 
 Dim intResponse As Integer 
 
 intResponse = MsgBox("Are you sure you want " & _ 
 "to delete the cells?", vbYesNo) 
 
 If intResponse = vbYes Then 
 ActiveDocument.Tables(1).Cell(1, 1).Delete 
 End If 
End Sub
```


## See also


[Cell Object](Word.Cell.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]