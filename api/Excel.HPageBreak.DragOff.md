---
title: HPageBreak.DragOff method (Excel)
keywords: vbaxl10.chm159075
f1_keywords:
- vbaxl10.chm159075
ms.prod: excel
api_name:
- Excel.HPageBreak.DragOff
ms.assetid: 80065224-c53d-3f45-8d94-c644502dac22
ms.date: 04/26/2019
localization_priority: Normal
---


# HPageBreak.DragOff method (Excel)

Drags a page break out of the print area.


## Syntax

_expression_.**DragOff** (_Direction_, _RegionIndex_)

_expression_ A variable that represents an **[HPageBreak](Excel.HPageBreak.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Direction_|Required| **[XlDirection](Excel.XlDirection.md)**|The direction in which the page break is dragged.|
| _RegionIndex_|Required| **Long**|The print-area region index for the page break (the region where the mouse pointer is located when the mouse button is pressed if the user drags the page break).<br/><br/>If the print area is contiguous, there's only one print region. If the print area is discontiguous, there's more than one print region.|

## Remarks

This method exists primarily for the macro recorder. You can use the **[Delete](Excel.HPageBreak.Delete.md)** method to delete a page break in Visual Basic.


## Example

This example deletes vertical page break one from the active sheet by dragging it off the right edge of print region one.

```vb
ActiveSheet.VPageBreaks(1).DragOff xlToRight, 1
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]