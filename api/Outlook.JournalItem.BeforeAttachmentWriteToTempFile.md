---
title: JournalItem.BeforeAttachmentWriteToTempFile event (Outlook)
ms.prod: outlook
api_name:
- Outlook.JournalItem.BeforeAttachmentWriteToTempFile
ms.assetid: 0564d2b2-a20a-9fd3-d942-59a97dc19992
ms.date: 06/08/2017
localization_priority: Normal
---


# JournalItem.BeforeAttachmentWriteToTempFile event (Outlook)

Occurs before an attachment associated with an instance of the parent object is written to a temporary file.


## Syntax

_expression_. `BeforeAttachmentWriteToTempFile`( `_Attachment_` , `_Cancel_` )

_expression_ A variable that represents a [JournalItem](Outlook.JournalItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](Outlook.Attachment.md)**|The  **Attachment** to be written.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be written.|

## See also


[JournalItem Object](Outlook.JournalItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]