---
title: Conflicts.Session property (Outlook)
keywords: vbaol11.chm402
f1_keywords:
- vbaol11.chm402
ms.prod: outlook
api_name:
- Outlook.Conflicts.Session
ms.assetid: 4f707a23-5687-7076-9297-3fc14c98731a
ms.date: 06/08/2017
localization_priority: Normal
---


# Conflicts.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [Conflicts](Outlook.Conflicts.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Conflicts Object](Outlook.Conflicts.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]