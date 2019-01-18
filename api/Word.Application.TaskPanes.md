---
title: Application.TaskPanes property (Word)
keywords: vbawd10.chm158335433
f1_keywords:
- vbawd10.chm158335433
ms.prod: word
api_name:
- Word.Application.TaskPanes
ms.assetid: 0b0add9d-6c76-9dca-e7a5-3f653f5d1581
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.TaskPanes property (Word)

Returns a  **[TaskPanes](Word.TaskPanes.md)** collection that represents the most commonly performed tasks in Microsoft Word.


## Syntax

 _expression_. `TaskPanes`

 _expression_ An expression that returns an '[Application](Word.Application.md)' object.


## Example

The following example displays the task pane that contains information about formatting in a document.


```vb
Sub showFormatting() 
 Application.TaskPanes.Item(wdTaskPaneFormatting).Visible = True 
End Sub
```


## See also


[Application Object](Word.Application.md)

