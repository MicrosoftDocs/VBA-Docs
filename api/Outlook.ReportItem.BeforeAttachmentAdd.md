---
title: ReportItem.BeforeAttachmentAdd event (Outlook)
ms.prod: outlook
api_name:
- Outlook.ReportItem.BeforeAttachmentAdd
ms.assetid: c8b45b3b-627c-4851-b743-2612828546b0
ms.date: 06/08/2017
localization_priority: Normal
---


# ReportItem.BeforeAttachmentAdd event (Outlook)

Occurs before an attachment is added to an instance of the parent object.


## Syntax

_expression_. `BeforeAttachmentAdd`( `_Attachment_` , `_Cancel_` )

_expression_ A variable that represents a [ReportItem](Outlook.ReportItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](Outlook.Attachment.md)**|The  **Attachment** to be added to the item.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be added.|

## See also


[ReportItem Object](Outlook.ReportItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]