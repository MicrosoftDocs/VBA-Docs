---
title: TasksModule.Session property (Outlook)
keywords: vbaol11.chm2844
f1_keywords:
- vbaol11.chm2844
ms.prod: outlook
api_name:
- Outlook.TasksModule.Session
ms.assetid: 947b6795-21db-e2fb-b76b-43dc90520403
ms.date: 06/08/2017
localization_priority: Normal
---


# TasksModule.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [TasksModule](Outlook.TasksModule.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[TasksModule Object](Outlook.TasksModule.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]