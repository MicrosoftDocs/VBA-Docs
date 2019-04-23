---
title: Reminder.Session property (Outlook)
keywords: vbaol11.chm556
f1_keywords:
- vbaol11.chm556
ms.prod: outlook
api_name:
- Outlook.Reminder.Session
ms.assetid: 30bd8c36-1afa-aae1-f050-47ad43af53f9
ms.date: 06/08/2017
localization_priority: Normal
---


# Reminder.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [Reminder](Outlook.Reminder.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Reminder Object](Outlook.Reminder.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]