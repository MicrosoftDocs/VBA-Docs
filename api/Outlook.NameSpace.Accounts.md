---
title: NameSpace.Accounts property (Outlook)
keywords: vbaol11.chm778
f1_keywords:
- vbaol11.chm778
ms.prod: outlook
api_name:
- Outlook.NameSpace.Accounts
ms.assetid: 80e969ea-d2cc-966d-5fe4-68d59951b5c9
ms.date: 06/08/2017
localization_priority: Normal
---


# NameSpace.Accounts property (Outlook)

Returns an **[Accounts](Outlook.Accounts.md)** collection object that represents all the **[Account](Outlook.Account.md)** objects in the current profile. Read-only.


## Syntax

_expression_. `Accounts`

_expression_ A variable that represents a [NameSpace](Outlook.NameSpace.md) object.


## Remarks

If Outlook is running in sessionless mode,  **Accounts** returns an **Accounts** collection with **[Accounts.Count](Outlook.Accounts.Count.md)** equal to 0.


## See also


[NameSpace Object](Outlook.NameSpace.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]