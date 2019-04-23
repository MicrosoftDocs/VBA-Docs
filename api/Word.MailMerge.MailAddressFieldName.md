---
title: MailMerge.MailAddressFieldName property (Word)
keywords: vbawd10.chm153092105
f1_keywords:
- vbawd10.chm153092105
ms.prod: word
api_name:
- Word.MailMerge.MailAddressFieldName
ms.assetid: 729e6afa-26a6-75dd-78f8-9677aedfb2fa
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMerge.MailAddressFieldName property (Word)

Returns or sets the name of the field that contains email addresses that are used when the mail merge destination is electronic mail. Read/write  **String**.


## Syntax

_expression_. `MailAddressFieldName`

 _expression_ An expression that returns a '[MailMerge](Word.MailMerge.md)' object.


## Example

This example merges the document named "FormLetter.doc" with its attached data document and sends the results to the email addresses stored in the Email merge field.


```vb
With Documents("FormLetter.doc").MailMerge 
 .MailAddressFieldName = "Email" 
 .MailSubject = "Amazing offer" 
 .Destination = wdSendToEmail 
 .Execute 
End With
```


## See also


[MailMerge Object](Word.MailMerge.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]