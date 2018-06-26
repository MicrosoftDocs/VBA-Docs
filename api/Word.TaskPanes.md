---
title: TaskPanes Object (Word)
ms.prod: word
api_name:
- Word.TaskPanes
ms.assetid: a560a41b-a1d7-175a-b475-af742c9fa1f8
ms.date: 06/08/2017
---


# TaskPanes Object (Word)

A collection of  **TaskPane** objects that contains commonly performed tasks in Microsoft Word.


## Remarks

Use the  **TaskPanes** property to return the **TaskPanes** collection. Use the **Item** method with a **[WdTaskPanes](Word.WdTaskPanes.md)** constant to refer to a specific task pane. The example below displays the formatting task pane.


```vb
Sub FormattingPane() 
 Application.TaskPanes(wdTaskPaneFormatting).Visible = True 
End Sub
```


## See also


[Word Object Model Reference](./overview/object-model-word-vba-reference.md)


