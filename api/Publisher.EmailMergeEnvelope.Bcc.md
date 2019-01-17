---
title: EmailMergeEnvelope.Bcc Property (Publisher)
keywords: vbapb10.chm9043974
f1_keywords:
- vbapb10.chm9043974
ms.prod: publisher
api_name:
- Publisher.EmailMergeEnvelope.Bcc
ms.assetid: 1d846fac-d93c-6a20-ce3b-090525dbbfe1
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailMergeEnvelope.Bcc Property (Publisher)

Gets or sets a semicolon-delimited list of email addresses that receive a blind carbon copy (BCC) of the email message. Read/write.


## Syntax

 _expression_. **Bcc**

 _expression_ A variable that represents an  **EmailMergeEnvelope** object.


## Return value

String


## Remarks

Set the  **Bcc** property to a string of email addresses separated by semicolons, as shown in the following example.


```vb
 MailMerge.EmailMergeEnvelope.Bcc = "name1@address1;name2@address2;name3@address3;..."
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]