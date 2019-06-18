---
title: Window.WindowState property (Publisher)
keywords: vbapb10.chm262160
f1_keywords:
- vbapb10.chm262160
ms.prod: publisher
api_name:
- Publisher.Window.WindowState
ms.assetid: 063ede5e-f279-09e3-5672-b634c752b927
ms.date: 06/18/2019
localization_priority: Normal
---


# Window.WindowState property (Publisher)

Returns or sets a **[PbWindowState](publisher.pbwindowstate.md)** constant indicating the state of the Microsoft Publisher window. Read/write.


## Syntax

_expression_.**WindowState**

_expression_ A variable that represents a **[Window](Publisher.Window.md)** object.


## Return value

PbWindowState


## Remarks

The **WindowState** property value can be one of the **PbWindowState** constants.

When the state of the window is **pbWindowStateNormal**, the window is neither maximized nor minimized.


## Example

This example maximizes the Publisher window.

```vb
ActiveWindow.WindowState = pbWindowStateMaximize
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]