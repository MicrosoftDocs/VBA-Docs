---
title: TaskRequestUpdateItem.Session property (Outlook)
keywords: vbaol11.chm1919
f1_keywords:
- vbaol11.chm1919
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem.Session
ms.assetid: 12e7fa2c-1067-4faa-c827-b1b1f8dc4238
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestUpdateItem.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [TaskRequestUpdateItem](Outlook.TaskRequestUpdateItem.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[TaskRequestUpdateItem Object](Outlook.TaskRequestUpdateItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]