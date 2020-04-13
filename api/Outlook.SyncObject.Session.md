---
title: SyncObject.Session property (Outlook)
keywords: vbaol11.chm105
f1_keywords:
- vbaol11.chm105
ms.prod: outlook
api_name:
- Outlook.SyncObject.Session
ms.assetid: 985369af-2fc0-8abd-d1c0-1fbb100a244d
ms.date: 06/08/2017
localization_priority: Normal
---


# SyncObject.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [SyncObject](Outlook.SyncObject.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[SyncObject Object](Outlook.SyncObject.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]