---
title: PostItem.BeforeAttachmentAdd event (Outlook)
ms.prod: outlook
api_name:
- Outlook.PostItem.BeforeAttachmentAdd
ms.assetid: 5617ce9e-0393-c122-3b85-dac92cb38f4b
ms.date: 06/08/2017
localization_priority: Normal
---


# PostItem.BeforeAttachmentAdd event (Outlook)

Occurs before an attachment is added to an instance of the parent object.


## Syntax

_expression_. `BeforeAttachmentAdd`( `_Attachment_` , `_Cancel_` )

_expression_ A variable that represents a [PostItem](Outlook.PostItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](Outlook.Attachment.md)**|The **Attachment** to be added to the item.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be added.|

## See also


[PostItem Object](Outlook.PostItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]