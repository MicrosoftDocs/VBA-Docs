---
title: UserProperty.Session property (Outlook)
keywords: vbaol11.chm215
f1_keywords:
- vbaol11.chm215
ms.prod: outlook
api_name:
- Outlook.UserProperty.Session
ms.assetid: 181d0aad-9b03-9cce-b6dd-33a290d57ee9
ms.date: 06/08/2017
localization_priority: Normal
---


# UserProperty.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [UserProperty](Outlook.UserProperty.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[UserProperty Object](Outlook.UserProperty.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]