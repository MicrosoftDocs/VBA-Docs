---
title: PostItem.BeforeAttachmentWriteToTempFile event (Outlook)
ms.prod: outlook
api_name:
- Outlook.PostItem.BeforeAttachmentWriteToTempFile
ms.assetid: c05d420d-8abe-2539-c8e6-64372828ec5c
ms.date: 06/08/2017
localization_priority: Normal
---


# PostItem.BeforeAttachmentWriteToTempFile event (Outlook)

Occurs before an attachment associated with an instance of the parent object is written to a temporary file.


## Syntax

_expression_. `BeforeAttachmentWriteToTempFile`( `_Attachment_` , `_Cancel_` )

_expression_ A variable that represents a [PostItem](Outlook.PostItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](Outlook.Attachment.md)**|The  **Attachment** to be written.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be written.|

## See also


[PostItem Object](Outlook.PostItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]