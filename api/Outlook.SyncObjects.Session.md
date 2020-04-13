---
title: SyncObjects.Session property (Outlook)
keywords: vbaol11.chm97
f1_keywords:
- vbaol11.chm97
ms.prod: outlook
api_name:
- Outlook.SyncObjects.Session
ms.assetid: 443c2e6d-fda7-8230-b3b1-bd87cccafe23
ms.date: 06/08/2017
localization_priority: Normal
---


# SyncObjects.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [SyncObjects](Outlook.SyncObjects.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[SyncObjects Object](Outlook.SyncObjects.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]