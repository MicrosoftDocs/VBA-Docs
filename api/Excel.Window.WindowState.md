---
title: Window.WindowState property (Excel)
keywords: vbaxl10.chm356125
f1_keywords:
- vbaxl10.chm356125
api_name:
- Excel.Window.WindowState
ms.assetid: be51b777-1370-03a2-1e3b-a4a89205f6ca
ms.date: 05/21/2019
ms.localizationpriority: medium
---


# Window.WindowState property (Excel)

Returns or sets the state of the window. Read/write **[XlWindowState](Excel.XlWindowState.md)**.


## Syntax

_expression_.**WindowState**

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


## Example

This example maximizes the application window in Microsoft Excel.

```vb
Application.WindowState = xlMaximized
```

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
