---
title: UserDefinedProperties.Session property (Outlook)
keywords: vbaol11.chm584
f1_keywords:
- vbaol11.chm584
ms.prod: outlook
api_name:
- Outlook.UserDefinedProperties.Session
ms.assetid: 7fb72c53-bb2e-5c27-61e6-a7ac79726647
ms.date: 06/08/2017
localization_priority: Normal
---


# UserDefinedProperties.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [UserDefinedProperties](Outlook.UserDefinedProperties.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[UserDefinedProperties Object](Outlook.UserDefinedProperties.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]