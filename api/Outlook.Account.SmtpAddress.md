---
title: Account.SmtpAddress property (Outlook)
keywords: vbaol11.chm743
f1_keywords:
- vbaol11.chm743
ms.prod: outlook
api_name:
- Outlook.Account.SmtpAddress
ms.assetid: 443beb7a-0ada-8e86-69d7-63880033abca
ms.date: 06/08/2017
localization_priority: Normal
---


# Account.SmtpAddress property (Outlook)

Returns a **String** representing the Simple Mail Transfer Protocol (SMTP) address for the **[Account](Outlook.Account.md)**. Read-only.


## Syntax

_expression_. `SmtpAddress`

_expression_ A variable that represents an [Account](Outlook.Account.md) object.


## Remarks

The purpose of  **SmtpAddress** and **[Account.UserName](Outlook.Account.UserName.md)** is to provide an account-based context to determine identity.

If the account does not have an SMTP address,  **SmtpAddress** returns an empty string.


## See also

- [Send an email given the SMTP address of an account](../outlook/How-to/Items-Folders-and-Stores/send-an-e-mail-given-the-smtp-address-of-an-account-outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]