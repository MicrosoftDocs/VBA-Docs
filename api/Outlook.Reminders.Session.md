---
title: Reminders.Session property (Outlook)
keywords: vbaol11.chm568
f1_keywords:
- vbaol11.chm568
ms.prod: outlook
api_name:
- Outlook.Reminders.Session
ms.assetid: 000e69b8-fd8c-1bd2-4cda-659faf210711
ms.date: 06/08/2017
localization_priority: Normal
---


# Reminders.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [Reminders](Outlook.Reminders.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Reminders Object](Outlook.Reminders.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]