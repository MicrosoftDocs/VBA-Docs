---
title: Account.Session property (Outlook)
keywords: vbaol11.chm738
f1_keywords:
- vbaol11.chm738
ms.prod: outlook
api_name:
- Outlook.Account.Session
ms.assetid: 92890235-402c-80c8-10b7-7339f153134e
ms.date: 06/08/2017
localization_priority: Normal
---


# Account.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [Account](Outlook.Account.md) object.


## Remarks

Returns  **Null** (**Nothing** in Visual Basic) if there is no logged-on session.

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:




```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```




```vb
Set objSession = Application.Session
```


## See also


[Account Object](Outlook.Account.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]