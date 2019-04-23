---
title: Cells.Delete method (Word)
keywords: vbawd10.chm155844808
f1_keywords:
- vbawd10.chm155844808
ms.prod: word
api_name:
- Word.Cells.Delete
ms.assetid: 891c21b7-ef8d-9ba1-9408-6560dac146c7
ms.date: 06/08/2017
localization_priority: Normal
---


# Cells.Delete method (Word)

Deletes a table cell or cells and optionally controls how the remaining cells are shifted.


## Syntax

_expression_.**Delete**( `_ShiftCells_` )

_expression_ Required. A variable that represents a '[Cells](Word.cells.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ShiftCells_|Optional| **Variant**|The direction in which the remaining cells are to be shifted. Can be any  **[WdDeleteCells](Word.WdDeleteCells.md)** constant. If omitted, cells to the right of the last deleted cell are shifted left.|

## See also


[Cells Collection Object](Word.cells.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]