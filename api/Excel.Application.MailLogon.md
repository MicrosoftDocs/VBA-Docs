---
title: Application.MailLogon method (Excel)
keywords: vbaxl10.chm133158
f1_keywords:
- vbaxl10.chm133158
ms.prod: excel
api_name:
- Excel.Application.MailLogon
ms.assetid: 0a6c8752-739d-b996-1426-4d3021ea5323
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.MailLogon method (Excel)

Logs on to MAPI Mail or Microsoft Exchange and establishes a mail session. If Microsoft Mail isn't already running, you must use this method to establish a mail session before mail or document routing functions can be used.


## Syntax

_expression_.**MailLogon** (_Name_, _Password_, _DownloadNewMail_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Optional| **Variant**|The mail account name or Microsoft Exchange profile name. If this argument is omitted, the default mail account name is used.|
| _Password_|Optional| **Variant**|The mail account password. This argument is ignored in Microsoft Exchange.|
| _DownloadNewMail_|Optional| **Variant**| **True** to download new mail immediately.|

## Remarks

Microsoft Excel logs off from any mail sessions it previously established before attempting to establish the new session.

To piggyback on the system default mail session, omit both the name and password parameters.


## Example

This example logs on to the default mail account.

```vb
If IsNull(Application.MailSession) Then 
 Application.MailLogon 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]