---
title: MailItem.BeforeAttachmentWriteToTempFile Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MailItem.BeforeAttachmentWriteToTempFile
ms.assetid: fad940fa-3ab8-ac9c-0cc1-adc36c695af8
ms.date: 06/08/2017
---


# MailItem.BeforeAttachmentWriteToTempFile Event (Outlook)

Occurs before an attachment associated with an instance of the parent object is written to a temporary file.


## Syntax

 _expression_. `BeforeAttachmentWriteToTempFile`( `_Attachment_` , `_Cancel_` )

 _expression_ A variable that represents a [MailItem](./Outlook.MailItem.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](Outlook.Attachment.md)**|The  **Attachment** to be written.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be written.|

## See also


[MailItem Object](Outlook.MailItem.md)

