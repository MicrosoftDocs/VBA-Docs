---
title: Workbook.ReplyWithChanges method (Excel)
keywords: vbaxl10.chm199207
f1_keywords:
- vbaxl10.chm199207
ms.prod: excel
api_name:
- Excel.Workbook.ReplyWithChanges
ms.assetid: 60424d69-0062-aa5e-ea8f-4fb07086167a
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.ReplyWithChanges method (Excel)

Sends an email message to the author of a workbook that has been sent out for review, notifying them that a reviewer has completed review of the workbook.


## Syntax

_expression_.**ReplyWithChanges** (_ShowMessage_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ShowMessage_|Optional| **Variant**| **False** does not display the message. **True** displays the message.|

## Remarks

Use the **[SendForReview](Excel.Workbook.SendForReview.md)** method to start a collaborative review of a workbook. If the **ReplyWithChanges** method is executed on a workbook that is not part of a collaborative review cycle, the user will receive an error.


## Example

This example automatically sends a notification to the author of a review workbook indicating that a reviewer has completed a review, without first displaying the email message to the reviewer. This example assumes that the active workbook is part of a collaborative review cycle.

```vb
Sub ReplyMsg() 
 
 ActiveWorkbook.ReplyWithChanges ShowMessage:=False 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]