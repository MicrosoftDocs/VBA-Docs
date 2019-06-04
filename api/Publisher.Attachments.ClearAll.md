---
title: Attachments.ClearAll method (Publisher)
keywords: vbapb10.chm569350
f1_keywords:
- vbapb10.chm569350
ms.prod: publisher
api_name:
- Publisher.Attachments.ClearAll
ms.assetid: ae4e4c60-56cb-f97b-06f4-bd0d2abac4ee
ms.date: 06/05/2019
localization_priority: Normal
---


# Attachments.ClearAll method (Publisher)

Clears (deletes) all the **[Attachment](Publisher.Attachment.md)** objects in the parent **Attachments** collection of an email merge message.


## Syntax

_expression_.**ClearAll**

_expression_ A variable that represents an **[Attachments](Publisher.Attachments.md)** object.


## Remarks

To clear an individual attachment, use the **[Delete](Publisher.Attachment.Delete.md)** method of the specific **Attachment** object.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to clear all the attachments to the message in an email merge. The code prints the number of current attachments to the message in the Immediate window and then deletes all of the **Attachment** objects in the collection.

```vb
Public Sub ClearAll_Example() 
 
 Dim pubAttachments As Publisher.Attachments 
 
 Dim pubMailMerge As Publisher.MailMerge 
 Dim pubEmailMergeEnvelope As Publisher.EmailMergeEnvelope 
 
 Set pubMailMerge = ThisDocument.MailMerge 
 Set pubEmailMergeEnvelope = pubMailMerge.EmailMergeEnvelope 
 Set pubAttachments = pubEmailMergeEnvelope.Attachments 
 
 Debug.Print pubAttachments.Count 
 pubAttachments.ClearAll 
 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]