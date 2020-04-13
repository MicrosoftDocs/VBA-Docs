---
title: Application.IsTrusted property (Outlook)
keywords: vbaol11.chm733
f1_keywords:
- vbaol11.chm733
ms.prod: outlook
api_name:
- Outlook.Application.IsTrusted
ms.assetid: 4caeb41a-9cc3-1195-22a9-ad8eae12ce53
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.IsTrusted property (Outlook)

Returns a **Boolean** to indicate if an add-in or external caller is considered trusted by Outlook. Read-only


## Syntax

_expression_. `IsTrusted`

_expression_ A variable that represents an **[Application](Outlook.Application.md)** object.


## Remarks

For out-of-process callers that have instantiated the  **[Application](Outlook.Application.md)** object, **IsTrusted** always returns **False**. For Outlook add-ins, **IsTrusted** returns **True** if and only if the add-in is considered trusted by Outlook.


## See also


[Application Object](Outlook.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]