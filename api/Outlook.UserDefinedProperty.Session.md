---
title: UserDefinedProperty.Session property (Outlook)
keywords: vbaol11.chm3
f1_keywords:
- vbaol11.chm3
ms.prod: outlook
api_name:
- Outlook.UserDefinedProperty.Session
ms.assetid: b47e79c1-e28c-48c8-f1cb-08844bf9716a
ms.date: 06/08/2017
localization_priority: Normal
---


# UserDefinedProperty.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [UserDefinedProperty](Outlook.UserDefinedProperty.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[UserDefinedProperty Object](Outlook.UserDefinedProperty.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]