---
title: Window.VerticalPercentScrolled property (Word)
keywords: vbawd10.chm157417496
f1_keywords:
- vbawd10.chm157417496
ms.prod: word
api_name:
- Word.Window.VerticalPercentScrolled
ms.assetid: 008d46d1-667a-9f32-1f8c-cb18ccde8a2f
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.VerticalPercentScrolled property (Word)

Returns or sets the vertical scroll position as a percentage of the document length. Read/write  **Long**.


## Syntax

_expression_. `VerticalPercentScrolled`

_expression_ Required. A variable that represents a **[Window](Word.Window.md)** object.


## Example

This example displays the percentage that the active window is scrolled vertically.


```vb
MsgBox ActiveDocument.ActiveWindow.VerticalPercentScrolled & "%"
```

This example scrolls the active window vertically by 10 percent.




```vb
Set aWindow = ActiveDocument.ActiveWindow 
aWindow.VerticalPercentScrolled = _ 
 aWindow.VerticalPercentScrolled + 10
```


## See also


[Window Object](Word.Window.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]