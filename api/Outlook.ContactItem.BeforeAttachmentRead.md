---
title: ContactItem.BeforeAttachmentRead event (Outlook)
ms.prod: outlook
api_name:
- Outlook.ContactItem.BeforeAttachmentRead
ms.assetid: ba862dea-f2e1-a864-f6c3-a8987c28bfcf
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.BeforeAttachmentRead event (Outlook)

Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an  **[Attachment](Outlook.Attachment.md)** object.


## Syntax

_expression_. `BeforeAttachmentRead`( `_Attachment_` , `_Cancel_` )

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **Attachment**|The  **Attachment** to be read.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be read.|

## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]