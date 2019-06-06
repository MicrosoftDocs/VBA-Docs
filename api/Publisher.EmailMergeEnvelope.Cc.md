---
title: EmailMergeEnvelope.Cc property (Publisher)
keywords: vbapb10.chm9043972
f1_keywords:
- vbapb10.chm9043972
ms.prod: publisher
api_name:
- Publisher.EmailMergeEnvelope.Cc
ms.assetid: d9e7704c-c45a-cf19-e0a8-8d55e1e82fc0
ms.date: 06/07/2019
localization_priority: Normal
---


# EmailMergeEnvelope.Cc property (Publisher)

Gets or sets the **[MailMergeDataField](publisher.mailmergedatafield.md)** object that represents the data-source field (column) that lists the email addresses of recipients that you want to receive a carbon copy (CC) of the merged email message. Read/write.


## Syntax

_expression_.**Cc**

_expression_ A variable that represents an **[EmailMergeEnvelope](Publisher.EmailMergeEnvelope.md)** object.


## Return value

MailMergeDataField


## Remarks

You must make certain that you assign the correct data-source field (the one that represents CC email addresses) to the **Cc** property. You can use the following line of code, which gets the value of the **Name** property of the **MailMergeDataField** object to which **Cc** is assigned, to ensure that you make the correct assignment.

```vb
Debug.Print ThisDocument.MailMerge.EmailMergeEnvelope.Cc.Name
```

<br/>

For an example of how to set the **Cc** property value, see the **[EmailMergeEnvelope](Publisher.EmailMergeEnvelope.md)** object.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]