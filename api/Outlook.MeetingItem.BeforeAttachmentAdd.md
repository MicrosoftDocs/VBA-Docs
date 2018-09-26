---
title: MeetingItem.BeforeAttachmentAdd Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MeetingItem.BeforeAttachmentAdd
ms.assetid: 9550ed34-0e04-eee0-b149-4df496c8e155
ms.date: 06/08/2017
---


# MeetingItem.BeforeAttachmentAdd Event (Outlook)

Occurs before an attachment is added to an instance of the parent object.


## Syntax

 _expression_. `BeforeAttachmentAdd`( `_Attachment_` , `_Cancel_` )

 _expression_ A variable that represents a [MeetingItem](./Outlook.MeetingItem.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Attachment_|Required| **[Attachment](Outlook.Attachment.md)**|The  **Attachment** to be added to the item.|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **Attachment** to be added.|

## See also


[MeetingItem Object](Outlook.MeetingItem.md)

