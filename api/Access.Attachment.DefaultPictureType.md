---
title: Attachment.DefaultPictureType property (Access)
keywords: vbaac10.chm10452
f1_keywords:
- vbaac10.chm10452
ms.prod: access
api_name:
- Access.Attachment.DefaultPictureType
ms.assetid: 77032908-5b98-7072-1e53-520485580746
ms.date: 02/07/2019
localization_priority: Normal
---


# Attachment.DefaultPictureType property (Access)

Gets or sets the method used to store the image specified by the **[DefaultPicture](Access.Attachment.DefaultPicture.md)** property in the database. Read/write **Byte**.


## Syntax

_expression_.**DefaultPictureType**

_expression_ A variable that represents an **[Attachment](Access.Attachment.md)** object.


## Remarks

The **DefaultPictureType** property uses the following settings.

|Setting|Value|Meaning|
|:-----|:-----|:-----|
|Embedded (Default)|0|The image is embedded with the specified **Attachment** control.|
|Linked|1|The image is stored outside of the database.|
|Shared|2|The image is added to the **[SharedResources](Access.SharedResources.md)** collection.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]