---
title: OrderFields.Session property (Outlook)
keywords: vbaol11.chm2674
f1_keywords:
- vbaol11.chm2674
ms.prod: outlook
api_name:
- Outlook.OrderFields.Session
ms.assetid: cf1ea6e2-a4fb-0d54-268a-fae589448129
ms.date: 06/08/2017
localization_priority: Normal
---


# OrderFields.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [OrderFields](Outlook.OrderFields.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[OrderFields Object](Outlook.OrderFields.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]