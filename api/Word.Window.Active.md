---
title: Window.Active property (Word)
keywords: vbawd10.chm157417498
f1_keywords:
- vbawd10.chm157417498
ms.prod: word
api_name:
- Word.Window.Active
ms.assetid: 8413477e-aee6-43c6-34e1-267a59718da3
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.Active property (Word)

 **True** if the specified window is active. Read-only **Boolean**.


## Syntax

_expression_.**Active**

_expression_ Required. A variable that represents a **[Window](Word.Window.md)** object.


## Example

This example activates the first window in the **Windows** collection, if the window isn't currently active.


```vb
Sub ActiveWin() 
 If Windows(1).Active = False Then Windows(1).Activate 
End Sub
```


## See also


[Window Object](Word.Window.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]