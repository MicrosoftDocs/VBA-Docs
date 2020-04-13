---
title: Recipients.Session property (Outlook)
keywords: vbaol11.chm228
f1_keywords:
- vbaol11.chm228
ms.prod: outlook
api_name:
- Outlook.Recipients.Session
ms.assetid: 41ddda3c-ca79-9387-b416-8334aeecc102
ms.date: 06/08/2017
localization_priority: Normal
---


# Recipients.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [Recipients](Outlook.Recipients.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Recipients Object](Outlook.Recipients.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]