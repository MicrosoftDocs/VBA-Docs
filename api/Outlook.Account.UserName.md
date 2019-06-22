---
title: Account.UserName property (Outlook)
keywords: vbaol11.chm742
f1_keywords:
- vbaol11.chm742
ms.prod: outlook
api_name:
- Outlook.Account.UserName
ms.assetid: 3ab96240-b68c-e2f7-83b9-6d6663c4880d
ms.date: 06/08/2017
localization_priority: Normal
---


# Account.UserName property (Outlook)

Returns a  **String** representing the user name for the **[Account](Outlook.Account.md)**. Read-only.


## Syntax

_expression_.**UserName**

_expression_ A variable that represents an [Account](Outlook.Account.md) object.


## Remarks

The purpose of  **[Account.SmtpAddress](Outlook.Account.SmtpAddress.md)** and **UserName** is to provide an account-based context to determine identity.

If the account does not have a user name defined,  **UserName** returns an empty string.


## See also


[Account Object](Outlook.Account.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]