---
title: Attachments.Add method (Publisher)
keywords: vbapb10.chm569349
f1_keywords:
- vbapb10.chm569349
ms.prod: publisher
api_name:
- Publisher.Attachments.Add
ms.assetid: dbf2eb67-5e28-a7e6-226f-feac9045186b
ms.date: 06/05/2019
localization_priority: Normal
---


# Attachments.Add method (Publisher)

Adds an **[Attachment](Publisher.Attachment.md)** object to the **Attachments** collection of a Microsoft Publisher publication.


## Syntax

_expression_.**Add** (_FileName_)

_expression_ A variable that represents an **[Attachments](Publisher.Attachments.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_FileName_|Required| **String**|File name of the attachment.|

## Return value

Attachment


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to add an attachment to the message in an email merge. The code adds an attachment to an email message and then prints the number of current attachments to the message in the Immediate window.

The attachment in this example is an image file at the root of the C drive. Before running the code, replace "_C:\image.jpg_" with the path to and name of the file on your computer that you want to add as an email attachment.

Before you can create an email merge, you must use the **[OpenDataSource](Publisher.MailMerge.OpenDataSource.md)** method of the **MailMerge** object to connect the active document to a data source. To run the merge, use the **[Execute](publisher.mailmerge.execute.md)** method of the **MailMerge** object. 

For an example of how to connect to a data source and create an email merge, see the **[EmailMergeEnvelope](Publisher.EmailMergeEnvelope.md)** object.


```vb
Public Sub Add_Example() 
 
 Dim pubAttachment As Publisher.Attachment 
 
 Set pubAttachment = ThisDocument.MailMerge.EmailMergeEnvelope.Attachments.Add("C:\image.jpg") 
 Debug.Print ThisDocument.MailMerge.EmailMergeEnvelope.Attachments.Count 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]