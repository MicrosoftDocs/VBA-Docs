---
title: TaskItem.ShowCategoriesDialog method (Outlook)
keywords: vbaol11.chm1760
f1_keywords:
- vbaol11.chm1760
ms.prod: outlook
api_name:
- Outlook.TaskItem.ShowCategoriesDialog
ms.assetid: f31b247b-1e8a-6ea8-3d66-cec400e87b70
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskItem.ShowCategoriesDialog method (Outlook)

Displays the  **Show Categories** dialog box, which allows you to select categories that correspond to the subject of the item.


## Syntax

_expression_. `ShowCategoriesDialog`

_expression_ A variable that represents a [TaskItem](Outlook.TaskItem.md) object.


## Example

The following Microsoft Visual Basic for Applications (VBA) example creates a new task item, displays the item on the screen, and opens up the  **Show Categories** dialog box.


```vb
Sub TaskItem() 
 
 'Creates a task item to access ShowCategoriesDialog 
 
 Dim olmyTaskItem As Outlook.TaskItem 
 
 'Create task item 
 
 Set olmyTaskItem = Application.CreateItem(olTaskItem) 
 
 
 
 olmyTaskItem.Subject = "Sales Reports" 
 
 'Display the item 
 
 olmyTaskItem.Display 
 
 'Display the Show categories dialog 
 
 olmyTaskItem.ShowCategoriesDialog 
 
End Sub
```


## See also


[TaskItem Object](Outlook.TaskItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]