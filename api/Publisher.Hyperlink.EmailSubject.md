---
title: Hyperlink.EmailSubject property (Publisher)
keywords: vbapb10.chm4587524
f1_keywords:
- vbapb10.chm4587524
ms.prod: publisher
api_name:
- Publisher.Hyperlink.EmailSubject
ms.assetid: 16b60648-56fe-b8ba-3424-0dd6e88727e6
ms.date: 06/08/2019
localization_priority: Normal
---


# Hyperlink.EmailSubject property (Publisher)

Sets or returns a **String** that represents the subject for email messages generated to process web form data. Read/write.


## Syntax

_expression_.**EmailSubject**

_expression_ A variable that represents a **[Hyperlink](Publisher.Hyperlink.md)** object.


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