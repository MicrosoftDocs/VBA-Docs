---
title: Application.WindowState property (Excel)
keywords: vbaxl10.chm133234
f1_keywords:
- vbaxl10.chm133234
ms.prod: excel
api_name:
- Excel.Application.WindowState
ms.assetid: f53d2bb8-b862-c55f-d9d5-68e705ca3415
ms.date: 06/08/2017
---


# Application.WindowState property (Excel)

Returns or sets the state of the window. Read/write  **[xlWindowState](Excel.XlWindowState.md)** .


## Syntax

 _expression_. `WindowState`

 _expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


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


## See also


[Application Object](Excel.Application(object).md)

