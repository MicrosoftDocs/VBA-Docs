---
title: Window.UsableWidth property (Excel)
keywords: vbaxl10.chm356120
f1_keywords:
- vbaxl10.chm356120
ms.prod: excel
api_name:
- Excel.Window.UsableWidth
ms.assetid: 7244a9e5-c4f0-715e-74c8-586101b368ce
ms.date: 05/21/2019
localization_priority: Normal
---


# Window.UsableWidth property (Excel)

Returns the maximum width of the space that a window can occupy in the application window area, in [points](../language/glossary/vbe-glossary.md#point). Read-only **Double**.


## Syntax

_expression_.**UsableWidth**

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


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