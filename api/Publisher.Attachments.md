---
title: Attachments Object (Publisher)
keywords: vbapb10.chm9175039
f1_keywords:
- vbapb10.chm9175039
ms.prod: publisher
api_name:
- Publisher.Attachments
ms.assetid: 61957961-8c75-992f-159c-51412ed309ea
ms.date: 06/08/2017
localization_priority: Normal
---


# Attachments Object (Publisher)

The collection of  **[Attachment](Publisher.Attachment.md)** objects that represents all the attachments to a merged email message.
 


## Remarks

The  **Attachments** collection corresponds to the list of attachments in the **Attachments** box in the **Merge to Email** dialog box in the Microsoft Publisher user interface (on the **File** menu, point to **Send Email**, click  **Send Email Merge**, and then click  **Options**).
 

 
To add an  **Attachment** object to the **Attachments** collection and thereby add an attachment to the list of attachments to the merged email that you want to send, use the **Attachments.Add** method.
 

 
To remove a single attachment from an email merge message, use the  **Attachment.Delete** method of the specific **Attachment** object that you want to remove from the **Attachments** collection.
 

 
To remove all the attachments to the merged email and thereby empty the  **Attachments** collection, use the **Attachments.ClearAll** method.
 

 
The default property of the  **Attachments** collection is the **Item** property.
 

 

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Add** method to add an attachment to an email merge message. The macro adds an **Attachment** object that represents a bitmap image to the **Attachments** collection of the active document. It also iterates through the **Attachments** collection and prints the name of each attachment in the **Immediate** window.
 

 
Before running this macro, place a file named  _image.bmp_ in the root of the C drive on your computer, or change the name and path of the file in the macro to specify the one you want to attach.
 

 
To send an email merge message, you must connect to a data source, create the email merge, and then send the message. For more information, see the **[EmailMergeEnvelope](Publisher.EmailMergeEnvelope.md)** object topic.
 

 



```vb
Public Sub Attachments_Example() 
 
 Dim pubAttachments As Publisher.Attachments 
 Dim pubAttachment As Publisher.Attachment 
 Dim pubAttachment_Added As Publisher.Attachment 
 Dim pubMailMerge As Publisher.MailMerge 
 Dim pubEmailMergeEnvelope As Publisher.EmailMergeEnvelope 
 
 Set pubMailMerge = ThisDocument.MailMerge 
 Set pubEmailMergeEnvelope = pubMailMerge.EmailMergeEnvelope 
 Set pubAttachments = pubEmailMergeEnvelope.Attachments 
 
 Set pubAttachment_Added = pubAttachments.Add("C:\image.bmp ") 
 
 For Each pubAttachment In pubAttachments 
 Debug.Print pubAttachment.Name 
 Next 
 
End Sub
```


## Methods



|Name|
|:-----|
|[Add](Publisher.Attachments.Add.md)|
|[ClearAll](Publisher.Attachments.ClearAll.md)|

## Properties



|Name|
|:-----|
|[Application](Publisher.Attachments.Application.md)|
|[Count](Publisher.Attachments.Count.md)|
|[Item](Publisher.Attachments.Item.md)|
|[Parent](Publisher.Attachments.Parent.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]