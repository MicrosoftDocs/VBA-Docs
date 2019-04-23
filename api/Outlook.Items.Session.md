---
title: Items.Session property (Outlook)
keywords: vbaol11.chm55
f1_keywords:
- vbaol11.chm55
ms.prod: outlook
api_name:
- Outlook.Items.Session
ms.assetid: 5c385dfc-042e-7649-0f32-5d34e53fca57
ms.date: 06/08/2017
localization_priority: Normal
---


# Items.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [Items](Outlook.Items.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Items Object](Outlook.Items.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]