---
title: DocumentItem.BeforeAttachmentWriteToTempFile event (Outlook)
ms.prod: outlook
api_name:
- Outlook.DocumentItem.BeforeAttachmentWriteToTempFile
ms.assetid: 09ec6f62-e5c6-1884-ba77-e4865978d0ba
ms.date: 06/08/2017
localization_priority: Normal
---


# DocumentItem.BeforeAttachmentWriteToTempFile event (Outlook)

Occurs before an attachment associated with an instance of the parent object is written to a temporary file.


## Syntax

_expression_. `BeforeAttachmentWriteToTempFile`( `_Attachment_` , `_Cancel_` )

_expression_ A variable that represents a [DocumentItem](Outlook.DocumentItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](Outlook.Attachment.md)**|The  **Attachment** to be written.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be written.|

## See also


[DocumentItem Object](Outlook.DocumentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]