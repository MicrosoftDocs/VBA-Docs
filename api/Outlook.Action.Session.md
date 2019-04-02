---
title: Action.Session property (Outlook)
keywords: vbaol11.chm12
f1_keywords:
- vbaol11.chm12
ms.prod: outlook
api_name:
- Outlook.Action.Session
ms.assetid: cfe619d2-3a7e-c8af-de17-be2363de0a56
ms.date: 06/08/2017
localization_priority: Normal
---


# Action.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [Action](Outlook.Action.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Action Object](Outlook.Action.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]