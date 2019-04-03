---
title: OrderField.Session property (Outlook)
keywords: vbaol11.chm2685
f1_keywords:
- vbaol11.chm2685
ms.prod: outlook
api_name:
- Outlook.OrderField.Session
ms.assetid: af14c535-9588-0e3a-b9ff-8a4c46d0fc25
ms.date: 06/08/2017
localization_priority: Normal
---


# OrderField.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [OrderField](Outlook.OrderField.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[OrderField Object](Outlook.OrderField.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]