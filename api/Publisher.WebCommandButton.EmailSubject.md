---
title: WebCommandButton.EmailSubject property (Publisher)
keywords: vbapb10.chm3932168
f1_keywords:
- vbapb10.chm3932168
ms.prod: publisher
api_name:
- Publisher.WebCommandButton.EmailSubject
ms.assetid: 4d29dacd-0da6-c706-515e-219daf5e349d
ms.date: 06/18/2019
localization_priority: Normal
---


# WebCommandButton.EmailSubject property (Publisher)

Sets or returns a **String** that represents the subject for email messages generated to process web form data. Read/write.


## Syntax

_expression_.**EmailSubject**

_expression_ A variable that represents a **[WebCommandButton](Publisher.WebCommandButton.md)** object.


## Example

This example sets Publisher to process data on the web form in the current publication by sending an email message with a subject line to a specified email address.

```vb
Sub WebFormData() 
 With ThisDocument.Pages(1).Shapes(1).WebCommandButton 
 .DataRetrievalMethod = pbSubmitDataRetrievalEmail 
 .EmailAddress = "someone@example.com" 
 .EmailSubject = "Web form data" 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]