---
title: Window.PointsToScreenPixelsX method (Excel)
keywords: vbaxl10.chm356129
f1_keywords:
- vbaxl10.chm356129
ms.prod: excel
api_name:
- Excel.Window.PointsToScreenPixelsX
ms.assetid: b637ae59-30fe-a5cd-2c0d-d9cb63c77d84
ms.date: 05/21/2019
localization_priority: Normal
---


# Window.PointsToScreenPixelsX method (Excel)

Converts a horizontal measurement from [points](../language/glossary/vbe-glossary.md#point) (document coordinates) to screen pixels (screen coordinates). Returns the converted measurement as a **Long** value.


## Syntax

_expression_.**PointsToScreenPixelsX** (_Points_)

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Points_|Required| **Long**|The number of points horizontally along the top of the document window, starting from the left.|

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