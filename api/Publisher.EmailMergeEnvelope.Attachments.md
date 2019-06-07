---
title: EmailMergeEnvelope.Attachments property (Publisher)
keywords: vbapb10.chm9043975
f1_keywords:
- vbapb10.chm9043975
ms.prod: publisher
api_name:
- Publisher.EmailMergeEnvelope.Attachments
ms.assetid: 53948bf7-2727-7b9c-a645-c9b954d5e023
ms.date: 06/07/2019
localization_priority: Normal
---


# EmailMergeEnvelope.Attachments property (Publisher)

Gets the list of a merged email message's attachments. Read-only.


## Syntax

_expression_.**Attachments**

_expression_ A variable that represents an **[EmailMergeEnvelope](Publisher.EmailMergeEnvelope.md)** object.


## Return value

Attachments


## Remarks

To add attachments to a merged email message, use the **[Add](Publisher.Attachments.Add.md)** method of the **Attachments** collection. 

To remove an attachment, use the **[Delete](Publisher.Attachment.Delete.md)** method of the **Attachment** object. 

To remove all attachments, use the **[ClearAll](Publisher.Attachments.ClearAll.md)** method of the **Attachments** collection.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]