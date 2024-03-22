---
title: Application.WindowState property (Excel)
keywords: vbaxl10.chm133234
f1_keywords:
- vbaxl10.chm133234
api_name:
- Excel.Application.WindowState
ms.assetid: f53d2bb8-b862-c55f-d9d5-68e705ca3415
ms.date: 04/05/2019
ms.localizationpriority: medium
---


# Application.WindowState property (Excel)

Returns or sets the state of the window. Read/write **[XlWindowState](Excel.XlWindowState.md)**.


## Syntax

_expression_.**WindowState**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


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
