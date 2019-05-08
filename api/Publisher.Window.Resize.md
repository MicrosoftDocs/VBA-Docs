---
title: Window.Resize method (Publisher)
keywords: vbapb10.chm262164
f1_keywords:
- vbapb10.chm262164
ms.prod: publisher
api_name:
- Publisher.Window.Resize
ms.assetid: 478e5f05-a2f9-c3b0-5dd0-3248272b2c37
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.Resize method (Publisher)

Sizes the Microsoft Publisher application window.


## Syntax

_expression_.**Resize**(**_Width_**,  **_Height_**)

_expression_ A variable that represents a  **Window** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|Width|Required| **Long**|The width of the window, in points.|
|Height|Required| **Long**|The height of the window, in points.|

## Remarks

If the window is maximized or minimized, an error occurs.

Use the  **[Width](Publisher.Window.Width.md)** and  **[Height](Publisher.Window.Height.md)** properties to set the window width and height independently.


## Example

This example resizes the Publisher application window to 7 inches wide by 6 inches high.


```vb
With Application.ActiveWindow 
 .WindowState = wdWindowStateNormal 
 .Resize Width:=InchesToPoints(7), Height:=InchesToPoints(6) 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]