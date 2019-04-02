---
title: ExchangeUser.Session property (Outlook)
keywords: vbaol11.chm2063
f1_keywords:
- vbaol11.chm2063
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.Session
ms.assetid: 7d2d23f0-c441-281a-1784-fe63dfa47b9f
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeUser.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [ExchangeUser](Outlook.ExchangeUser.md) object.


## Remarks

The  **Session** property and the **[Application.GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[ExchangeUser Object](Outlook.ExchangeUser.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]