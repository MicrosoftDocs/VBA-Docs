---
title: Window.Width property (Publisher)
keywords: vbapb10.chm262150
f1_keywords:
- vbapb10.chm262150
ms.prod: publisher
api_name:
- Publisher.Window.Width
ms.assetid: 762df30a-7fdd-8f95-f64b-eae57e7c02fe
ms.date: 06/18/2019
localization_priority: Normal
---


# Window.Width property (Publisher)

Returns or sets a **Long** that represents the width (in [points](../language/glossary/vbe-glossary.md#point)) of the window. Read/write.


## Syntax

_expression_.**Width**

_expression_ A variable that represents a **[Window](Publisher.Window.md)** object.


## Example

This example sets the height and width of the active window if the window is neither maximized nor minimized.

```vb
Sub SetWindowHeight() 
 With ActiveWindow 
 If .WindowState = pbWindowStateNormal Then 
 .Height = InchesToPoints(5) 
 .Width = InchesToPoints(5) 
 End If 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]