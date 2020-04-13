---
title: Task object (Word)
keywords: vbawd10.chm2434
f1_keywords:
- vbawd10.chm2434
ms.prod: word
api_name:
- Word.Task
ms.assetid: 8802fcd5-0947-2ea0-308a-376077633e34
ms.date: 06/08/2017
localization_priority: Normal
---


# Task object (Word)

Represents a single task running on the system. The **Task** object is a member of the **[Tasks](Word.tasks.md)** collection. The **Tasks** collection includes all the applications that are currently running on the system.


## Remarks

Use  **Tasks** (Index), where Index is the application name or the index number, to return a single **Task** object. The following example switches to and resizes the application window for the first visible task in the **Tasks** collection.


```vb
With Tasks(1) 
 If .Visible = True Then 
 .Activate 
 .Width = 400 
 .Height = 200 
 End If 
End With
```

The following example restores the Calculator application window if Calculator is in the **[Tasks](Word.tasks.md)** collection.




```vb
If Tasks.Exists("Calculator") = True Then 
 Tasks("Calculator").WindowState = wdWindowStateNormal 
End If
```

Use Visual Basic's  **Shell** function to run an executable program and add the program to the **[Tasks](Word.tasks.md)** collection.


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]