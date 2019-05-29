---
title: Worksheet.MailEnvelope property (Excel)
keywords: vbaxl10.chm175150
f1_keywords:
- vbaxl10.chm175150
ms.prod: excel
api_name:
- Excel.Worksheet.MailEnvelope
ms.assetid: 9490f86c-a82f-d1ab-7315-29b89c799301
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.MailEnvelope property (Excel)

Represents an email header for a document.


## Syntax

_expression_.**MailEnvelope**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Example

This example sets the comments for the header of the active worksheet.

```vb
Sub HeaderComments() 
 
 ActiveSheet.MailEnvelope.Introduction = "To Whom It May Concern: " 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]