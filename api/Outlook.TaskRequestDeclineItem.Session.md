---
title: TaskRequestDeclineItem.Session property (Outlook)
keywords: vbaol11.chm1821
f1_keywords:
- vbaol11.chm1821
ms.prod: outlook
api_name:
- Outlook.TaskRequestDeclineItem.Session
ms.assetid: ca771a84-1cc6-b1ef-2dbf-ed05541b96d5
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestDeclineItem.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [TaskRequestDeclineItem](Outlook.TaskRequestDeclineItem.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[TaskRequestDeclineItem Object](Outlook.TaskRequestDeclineItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]