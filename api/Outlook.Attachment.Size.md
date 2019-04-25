---
title: Attachment.Size property (Outlook)
keywords: vbaol11.chm2375
f1_keywords:
- vbaol11.chm2375
ms.prod: outlook
api_name:
- Outlook.Attachment.Size
ms.assetid: 7a300b59-3d58-c2d0-afa3-c3e7ef6450b7
ms.date: 06/08/2017
localization_priority: Normal
---


# Attachment.Size property (Outlook)

Returns a  **Long** indicating the size (in bytes) of the attachment. Read-only.


## Syntax

_expression_.**Size**

_expression_ A variable that represents an [Attachment](Outlook.Attachment.md) object.


## Remarks

This information may not always be available for attachments. For example, on an S/MIME message, the actual attachment sizes are unknown until the attachment is extracted. In cases where the attachment size cannot be determined, this property returns 0.


## See also


[Attachment Object](Outlook.Attachment.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]