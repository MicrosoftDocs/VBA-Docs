---
title: Window.Top property (Publisher)
keywords: vbapb10.chm262148
f1_keywords:
- vbapb10.chm262148
ms.prod: publisher
api_name:
- Publisher.Window.Top
ms.assetid: 22fe0170-7433-a917-87ca-f418c2aefc70
ms.date: 06/18/2019
localization_priority: Normal
---


# Window.Top property (Publisher)

Returns or sets a **Long** that represents the distance between the top edge of the screen and the application window. Read/write.


## Syntax

_expression_.**Top**

_expression_ A variable that represents a **[Window](Publisher.Window.md)** object.


## Example

This example verifies that the state of application window is neither maximized nor minimized, and then resizes the window and moves it to 150 points from the top of the screen.

```vb
Sub MoveWindow() 
 With ActiveWindow 
 If .WindowState = pbWindowStateNormal Then 
 .Top = 150 
 .Resize Width:=500, Height:=500 
 End If 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]