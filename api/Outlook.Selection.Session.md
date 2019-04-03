---
title: Selection.Session property (Outlook)
keywords: vbaol11.chm83
f1_keywords:
- vbaol11.chm83
ms.prod: outlook
api_name:
- Outlook.Selection.Session
ms.assetid: 22390a36-a51c-615d-a646-45e5aa7d253f
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [Selection](Outlook.Selection.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Selection Object](Outlook.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]