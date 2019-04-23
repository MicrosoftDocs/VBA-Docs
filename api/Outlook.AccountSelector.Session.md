---
title: AccountSelector.Session property (Outlook)
keywords: vbaol11.chm3451
f1_keywords:
- vbaol11.chm3451
ms.prod: outlook
api_name:
- Outlook.AccountSelector.Session
ms.assetid: aca5ce47-5597-8bb3-588b-0c58d704b158
ms.date: 06/08/2017
localization_priority: Normal
---


# AccountSelector.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an '[AccountSelector](Outlook.AccountSelector.md)' object.


## Remarks

Returns  **Null** (**Nothing** in Visual Basic) if there is no logged-on session.

You can use the  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method of the **[Application](Outlook.Application.md)** object interchangeably to obtain the **NameSpace** object for the current session.


## See also


[AccountSelector Object](Outlook.AccountSelector.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]