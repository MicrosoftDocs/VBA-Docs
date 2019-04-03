---
title: Conflict.Session property (Outlook)
keywords: vbaol11.chm413
f1_keywords:
- vbaol11.chm413
ms.prod: outlook
api_name:
- Outlook.Conflict.Session
ms.assetid: cd7eaf1e-545b-5a40-d95c-841f72a7a15e
ms.date: 06/08/2017
localization_priority: Normal
---


# Conflict.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [Conflict](Outlook.Conflict.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Conflict Object](Outlook.Conflict.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]