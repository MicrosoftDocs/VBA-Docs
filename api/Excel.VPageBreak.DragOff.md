---
title: VPageBreak.DragOff method (Excel)
keywords: vbaxl10.chm156075
f1_keywords:
- vbaxl10.chm156075
ms.prod: excel
api_name:
- Excel.VPageBreak.DragOff
ms.assetid: 93e169e8-e2d6-4cca-bd82-2d11fdc1ae4c
ms.date: 05/18/2019
localization_priority: Normal
---


# VPageBreak.DragOff method (Excel)

Drags a page break out of the print area.


## Syntax

_expression_.**DragOff** (_Direction_, _RegionIndex_)

_expression_ A variable that represents a **[VPageBreak](Excel.VPageBreak.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Direction_|Required| **[XlDirection](Excel.XlDirection.md)**|The direction in which the page break is dragged.|
| _RegionIndex_|Required| **Long**|The print-area region index for the page break (the region where the mouse pointer is located when the mouse button is pressed if the user drags the page break). If the print area is contiguous, there's only one print region. If the print area is discontiguous, there's more than one print region.|

## Remarks

This method exists primarily for the macro recorder. You can use the **[Delete](Excel.VPageBreak.Delete.md)** method to delete a page break in Visual Basic.


## Example

This example deletes vertical page-break one from the active sheet by dragging it off the right edge of print region one.

```vb
ActiveSheet.VPageBreaks(1).DragOff xlToRight, 1
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]