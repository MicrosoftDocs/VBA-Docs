---
title: Application.Width property (Excel)
keywords: vbaxl10.chm133232
f1_keywords:
- vbaxl10.chm133232
ms.prod: excel
api_name:
- Excel.Application.Width
ms.assetid: eeb8ff27-d219-bade-3e0b-aed6e34d17d7
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.Width property (Excel)

Returns or sets a **Double** value that represents the distance, in [points](../language/glossary/vbe-glossary.md#point), from the left edge of the application window to its right edge.


## Syntax

_expression_.**Width**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

If the window is minimized, **Width** is read-only and returns the width of the window icon.


## Example

This example expands the active window to the maximum size available (assuming that the window isn't maximized).

```vb
With ActiveWindow 
 .WindowState = xlNormal 
 .Top = 1 
 .Left = 1 
 .Height = Application.UsableHeight 
 .Width = Application.UsableWidth 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
