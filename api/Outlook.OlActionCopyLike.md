---
title: OlActionCopyLike enumeration (Outlook)
keywords: vbaol11.chm3048
f1_keywords:
- vbaol11.chm3048
ms.prod: outlook
api_name:
- Outlook.OlActionCopyLike
ms.assetid: f566bbb1-4906-c1c2-1f8e-9f1564e6c072
ms.date: 06/08/2017
localization_priority: Normal
---


# OlActionCopyLike enumeration (Outlook)

Specifies how item properties will be copied.



|Name|Value|Description|
|:-----|:-----|:-----|
| **olForward**|2|Properties of the new item will be set such that the new item is a forward of the original item. If creating a new **[MailItem](Outlook.MailItem.md)**, the value of the **To** and **CC** properties in the new item will be empty and the **Subject** property of the new item will be the original **Subject** with a prefix such as "FW:" (see **[Prefix](Outlook.Action.Prefix.md)** property) added. The attachments on the original item will be copied to the new item.|
| **olReply**|0|Properties of the new item will be set such that the new item is a reply to the original item. If creating a new **[MailItem](Outlook.MailItem.md)**, the value of the original **To** field will be copied to the **SenderEmailAddress** property of the new item, the **CC** field will be blank and the **Subject** field of the new item will be the original **Subject** with a prefix such as "RE:" (see **[Prefix](Outlook.Action.Prefix.md)** property) added.|
| **olReplyAll**|1|Properties of the new item will be set such that the new item is a reply to all of the senders of the original item. If creating a new **[MailItem](Outlook.MailItem.md)**, the value of the **SenderEmailAddress** and **CC** properties will be copied to the **To** property of the new item and the **Subject** property of the new item will be the Subject of the original item with a prefix such as "RE:" (see **[Prefix](Outlook.Action.Prefix.md)** property) added.|
| **olReplyFolder**|3|If creating a new **[PostItem](Outlook.PostItem.md)** based on an old one, the Post To property of the new item will contain the active folder address, the **Subject** property of the original item will be copied to the **ConversationTopic** property of the new item, and the **Subject** property of the new item will be empty.|
| **olRespond**|4|Used exclusively for voting button actions.|

## Remarks

Used by the [CopyLike](Outlook.Action.CopyLike.md) property of an [Action](Outlook.Action.md) to specify how item properties will be copied to the new item that is created when the **Action** is executed.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]