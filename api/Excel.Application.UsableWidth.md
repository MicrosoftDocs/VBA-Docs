---
title: Application.UsableWidth property (Excel)
keywords: vbaxl10.chm133223
f1_keywords:
- vbaxl10.chm133223
ms.prod: excel
api_name:
- Excel.Application.UsableWidth
ms.assetid: b6c1cecb-28a5-8cdf-95ae-1b3b6e200dbb
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.UsableWidth property (Excel)

Returns the maximum width of the space that a window can occupy in the application window area, in points. Read-only **Double**.


## Syntax

_expression_.**UsableWidth**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example expands the active window to the maximum size available (assuming that the window isn't already maximized).

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