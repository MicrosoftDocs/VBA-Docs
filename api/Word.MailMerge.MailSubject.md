---
title: MailMerge.MailSubject property (Word)
keywords: vbawd10.chm153092106
f1_keywords:
- vbawd10.chm153092106
ms.prod: word
api_name:
- Word.MailMerge.MailSubject
ms.assetid: 75303fd3-5d9f-e790-8ade-a7433c451a66
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMerge.MailSubject property (Word)

Returns or sets the subject line used when the mail merge destination is electronic mail. Read/write  **String**.


## Syntax

_expression_. `MailSubject`

 _expression_ An expression that returns a '[MailMerge](Word.MailMerge.md)' object.


## Example

This example merges the document named "Offer.doc" with its attached data document. The results are sent to the email addresses stored in the EmailNames merge field, and the subject of the mail message is "Amazing Offer."


```vb
With Documents("Offer.doc").MailMerge 
 .MailAddressFieldName = "EmailNames" 
 .MailSubject = "Amazing Offer" 
 .Destination = wdSendToEmail 
 .Execute 
End With
```


## See also


[MailMerge Object](Word.MailMerge.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]