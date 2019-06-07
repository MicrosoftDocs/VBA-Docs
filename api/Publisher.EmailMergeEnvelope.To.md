---
title: EmailMergeEnvelope.To property (Publisher)
keywords: vbapb10.chm9043971
f1_keywords:
- vbapb10.chm9043971
ms.prod: publisher
api_name:
- Publisher.EmailMergeEnvelope.To
ms.assetid: c9c470e8-1411-fda9-becf-5c932e97d98f
ms.date: 06/07/2019
localization_priority: Normal
---


# EmailMergeEnvelope.To property (Publisher)

Gets or sets the **[MailMergeDataField](publisher.mailmergedatafield.md)** object that represents the data-source field (column) that lists the email addresses of recipients of the merged email message. Read/write.


## Syntax

_expression_.**To**

_expression_ A variable that represents an **[EmailMergeEnvelope](Publisher.EmailMergeEnvelope.md)** object.


## Return value

MailMergeDataField


## Remarks

You must make certain that you assign the correct data-source field (the one that represents email addresses) to the **To** property. You can use the following line of code, which gets the value of the **Name** property of the **MailMergeDataField** object to which **To** is assigned, to ensure that you make the correct assignment.

```vb
Debug.Print ThisDocument.MailMerge.EmailMergeEnvelope.To.Name
```

<br/>

For an example of how to set the **To** property value, see the **[EmailMergeEnvelope](Publisher.EmailMergeEnvelope.md)** object.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]