---
title: Workbook.SendForReview method (Excel)
keywords: vbaxl10.chm199206
f1_keywords:
- vbaxl10.chm199206
ms.prod: excel
api_name:
- Excel.Workbook.SendForReview
ms.assetid: 3834f5b3-6d24-1bb9-27b5-052aa2e725e3
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.SendForReview method (Excel)

Sends a workbook in an email message for review to the specified recipients.


## Syntax

_expression_.**SendForReview** (_Recipients_, _Subject_, _ShowMessage_, _IncludeAttachment_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Recipients_|Optional| **Variant**|A string that lists the people to whom to send the message. These can be unresolved names and aliases in an email phone book or full email addresses. Separate multiple recipients with a semicolon (;). If left blank and  _ShowMessage_ is **False**, you will receive an error message, and the message will not be sent.|
| _Subject_|Optional| **Variant**|A string for the subject of the message. If left blank, the subject will be: Please review "filename".|
| _ShowMessage_|Optional| **Variant**|A **Boolean** value that indicates whether the message should be displayed when the method is executed. The default value is **True**. If set to **False**, the message is automatically sent to the recipients without first showing the message to the sender.|
| _IncludeAttachment_|Optional| **Variant**|A **Boolean** value that indicates whether the message should include an attachment or a link to a server location. The default value is **True**. If set to **False**, the document must be stored at a shared location.|

## Remarks

The **SendForReview** method starts a collaborative review cycle. Use the **[EndReview](Excel.Workbook.EndReview.md)** method to end a review cycle.


## Example

This example automatically sends the active workbook as an attachment in an email message to the specified recipients.

```vb
Sub WebReview() 
 
 ActiveWorkbook.SendForReview _ 
 Recipients:="someone@example.com; amy jones; lewjudy", _ 
 Subject:="Please review this document.", _ 
 ShowMessage:=False, _ 
 IncludeAttachment:=True 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]