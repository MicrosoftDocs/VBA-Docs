---
title: Window.DisplayVerticalRuler property (Word)
keywords: vbawd10.chm157417485
f1_keywords:
- vbawd10.chm157417485
ms.prod: word
api_name:
- Word.Window.DisplayVerticalRuler
ms.assetid: a529b86a-80a1-0ee3-821c-f11bcdb2a9ca
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.DisplayVerticalRuler property (Word)

 **True** if a vertical ruler is displayed for the specified window or pane. Read/write **Boolean**.


## Syntax

_expression_. `DisplayVerticalRuler`

_expression_ A variable that represents a **[Window](Word.Window.md)** object.


## Remarks

A vertical ruler appears only in print layout view, and only if the **DisplayRulers** property is set to **True**.


## Example

This example switches each window in the **Windows** collection to print layout view and displays the horizontal and vertical rulers.


```vb
Dim winLoop As Window 
 
For Each winLoop In Windows 
 With winLoop 
 .View.Type = wdPrintView 
 .DisplayRulers = True 
 .DisplayVerticalRuler = True 
 End With 
Next winLoop
```

This example hides the horizontal and vertical rulers for the active window.




```vb
With ActiveDocument.ActiveWindow 
 .DisplayVerticalRuler = False 
 .DisplayRulers = False 
End With
```


## See also


[Window Object](Word.Window.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]