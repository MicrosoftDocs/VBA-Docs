---
title: Application.MailSystem property (Excel)
keywords: vbaxl10.chm133160
f1_keywords:
- vbaxl10.chm133160
ms.prod: excel
api_name:
- Excel.Application.MailSystem
ms.assetid: df7b1238-bdf5-d9f8-9f50-585b489fd8a8
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.MailSystem property (Excel)

Returns the mail system that's installed on the host machine. Read-only **[XlMailSystem](Excel.XlMailSystem.md)**.


## Syntax

_expression_.**MailSystem**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

**xlMailSystem** can be one of these constants:

- **xlNoMailSystem**
- **xlPowerTalk**

## Example

This example displays the name of the mail system that's installed on the computer.

```vb
Select Case Application.MailSystem 
 Case xlMAPI 
 MsgBox "Mail system is Microsoft Mail" 
 Case xlPowerTalk 
 MsgBox "Mail system is PowerTalk" 
 Case xlNoMailSystem 
 MsgBox "No mail system installed" 
End Select
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]