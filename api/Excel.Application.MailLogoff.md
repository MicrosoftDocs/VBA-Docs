---
title: Application.MailLogoff method (Excel)
keywords: vbaxl10.chm133157
f1_keywords:
- vbaxl10.chm133157
ms.prod: excel
api_name:
- Excel.Application.MailLogoff
ms.assetid: 5265e9c1-6c04-3591-7133-5274e5b56347
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.MailLogoff method (Excel)

Closes a MAPI mail session established by Microsoft Excel.


## Syntax

_expression_.**MailLogoff**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

You cannot use this method to close or log off from Microsoft Mail.


## Example

This example closes the established mail session if there is one.

```vb
If Not IsNull(Application.MailSession) Then Application.MailLogoff
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]