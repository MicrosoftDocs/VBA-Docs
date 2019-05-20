---
title: Window.PointsToScreenPixelsY method (Excel)
keywords: vbaxl10.chm356130
f1_keywords:
- vbaxl10.chm356130
ms.prod: excel
api_name:
- Excel.Window.PointsToScreenPixelsY
ms.assetid: ec25e6d4-22c1-2444-9582-37187901ae02
ms.date: 05/21/2019
localization_priority: Normal
---


# Window.PointsToScreenPixelsY method (Excel)

Converts a vertical measurement from [points](../language/glossary/vbe-glossary.md#point) (document coordinates) to screen pixels (screen coordinates). Returns the converted measurement as a **Long** value.


## Syntax

_expression_.**PointsToScreenPixelsY** (_Points_)

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Points_|Required| **Long**|The number of points vertically along the left edge of the document window, starting from the top.|

## Return value

Long


## Example

This example determines the height and width (in pixels) of the selected cells in the active window and returns the values in the `lWinWidth` and `lWinHeight` variables.

```vb
With ActiveWindow 
 lWinWidth = _ 
 .PointsToScreenPixelsX(.Selection.Width) 
 lWinHeight = _ 
 .PointsToScreenPixelsY(.Selection.Height) 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]