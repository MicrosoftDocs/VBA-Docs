---
title: WebCommandButton.EmailAddress property (Publisher)
keywords: vbapb10.chm3932167
f1_keywords:
- vbapb10.chm3932167
ms.prod: publisher
api_name:
- Publisher.WebCommandButton.EmailAddress
ms.assetid: 8961e459-1ce1-558a-2450-c3b8da2d5559
ms.date: 06/18/2019
localization_priority: Normal
---


# WebCommandButton.EmailAddress property (Publisher)

Sets or returns a **String** representing the email address to use when processing web form data. Read/write.


## Syntax

_expression_.**EmailAddress**

_expression_ A variable that represents a **[WebCommandButton](Publisher.WebCommandButton.md)** object.


## Return value

String


## Example

This example sets Microsoft Publisher to process data on the web form in the current publication by sending an email message to a specified email address.

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