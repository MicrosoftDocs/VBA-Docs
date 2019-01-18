---
title: Attachment Object (Publisher)
keywords: vbapb10.chm9240575
f1_keywords:
- vbapb10.chm9240575
ms.prod: publisher
api_name:
- Publisher.Attachment
ms.assetid: d617bdf6-b0ba-be0d-0f72-f729010636c1
ms.date: 06/08/2017
localization_priority: Normal
---


# Attachment Object (Publisher)

Represents an attachment to a merged email message.


## Remarks

An **Attachment** object corresponds to one of the attachments in the list of attachments in the **Attachments** box in the **Merge to Email** dialog box in the Microsoft Publisher user interface. (On the **File** menu, point to **Send Email**, click  **Send Email Merge**, and then click  **Options**.)

To remove the attachment from the merged email, use the  **Delete** method of the **Attachment** object.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Add** method to add an attachment to an email merge message. It adds an **Attachment** object that represents a bitmap image to the **Attachments** collection of the active document.

Before running this macro, place a file named  _image.bmp_ in the root of the C drive on your computer, or change the name and file path of the file in the macro to specify the one you want to attach.

Note that to send an email merge message, you must connect to a data source, create the email merge, and then send the message. For more information, see the  **[EmailMergeEnvelope](./Publisher.EmailMergeEnvelope.md)** object topic.




```vb
Public Sub Attachment_Example() 
 
 Dim pubAttachments As Publisher.Attachments 
 Dim pubAttachment As Publisher.Attachment 
 Dim pubMailMerge As Publisher.MailMerge 
 Dim pubEmailMergeEnvelope As Publisher.EmailMergeEnvelope 
 
 Set pubMailMerge = ThisDocument.MailMerge 
 Set pubEmailMergeEnvelope = pubMailMerge.EmailMergeEnvelope 
 Set pubAttachments = pubEmailMergeEnvelope.Attachments 
 
 Set pubAttachment = pubAttachments.Add("C:\image.bmp ") 
 
End Sub
```


## Methods



|Name|
|:-----|
|[Delete](./Publisher.Attachment.Delete.md)|

## Properties



|Name|
|:-----|
|[Name](./Publisher.Attachment.Name.md)|

## See also


[Attachment Object Members](./overview/Publisher.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]