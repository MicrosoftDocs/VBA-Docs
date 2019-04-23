---
title: Application.MailSession property (Excel)
keywords: vbaxl10.chm133159
f1_keywords:
- vbaxl10.chm133159
ms.prod: excel
api_name:
- Excel.Application.MailSession
ms.assetid: 45dbbaa1-3da2-55f9-415b-ac9218d293dc
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.MailSession property (Excel)

Returns the MAPI mail session number as a hexadecimal string (if there's an active session), or returns **null** if there's no session. Read-only **Variant**.


## Syntax

_expression_.**MailSession**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

This property applies only to mail sessions created by Microsoft Excel (it doesn't return a mail session number for Microsoft Mail).

This property isn't used on PowerTalk mail systems.


## Example

This example closes the established mail session if there is one.

```vb
If Not IsNull(Application.MailSession) Then Application.MailLogoff
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]