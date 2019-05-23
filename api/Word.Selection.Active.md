---
title: Selection.Active property (Word)
keywords: vbawd10.chm158663059
f1_keywords:
- vbawd10.chm158663059
ms.prod: word
api_name:
- Word.Selection.Active
ms.assetid: a279837e-8ae7-24ec-71f0-de82c5a33ad8
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Active property (Word)

 **True** if the selection in the specified window or pane is active. Read-only **Boolean**.


## Syntax

_expression_.**Active**

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Example

This example splits the active window into two panes and activates the selection in the first pane, if it isn't already active.


```vb
Sub SplitWindow() 
 ActiveDocument.ActiveWindow.Split = True 
 If ActiveDocument.ActiveWindow.Panes(1).Selection _ 
 .Active = False Then 
 ActiveDocument.ActiveWindow.Panes(1).Activate 
 End If 
End Sub
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]