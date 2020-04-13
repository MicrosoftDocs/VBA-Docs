---
title: NameSpace.Session property (Outlook)
keywords: vbaol11.chm754
f1_keywords:
- vbaol11.chm754
ms.prod: outlook
api_name:
- Outlook.NameSpace.Session
ms.assetid: 93dba2e5-d11e-7761-ac29-08f5b7a83b49
ms.date: 06/08/2017
localization_priority: Normal
---


# NameSpace.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [NameSpace](Outlook.NameSpace.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[NameSpace Object](Outlook.NameSpace.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]