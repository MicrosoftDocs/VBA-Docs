---
title: Workbook.SendMail method (Excel)
keywords: vbaxl10.chm199149
f1_keywords:
- vbaxl10.chm199149
ms.prod: excel
api_name:
- Excel.Workbook.SendMail
ms.assetid: 581d197c-0748-2225-2986-64aa368aab39
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.SendMail method (Excel)

Sends the workbook by using the installed mail system.


## Syntax

_expression_.**SendMail** (_Recipients_, _Subject_, _ReturnReceipt_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Recipients_|Required| **Variant**|Specifies the name of the recipient as text, or as an array of text strings if there are multiple recipients. At least one recipient must be specified, and all recipients are added as To recipients.|
| _Subject_|Optional| **Variant**|Specifies the subject of the message. If this argument is omitted, the document name is used.|
| _ReturnReceipt_|Optional| **Variant**| **True** to request a return receipt. **False** to not request a return receipt. The default value is **False**.|

## Example

This example sends the active workbook to a single recipient.

```vb
ActiveWorkbook.SendMail recipients:="Jean Selva"
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
