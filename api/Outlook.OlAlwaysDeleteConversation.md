---
title: OlAlwaysDeleteConversation enumeration (Outlook)
keywords: vbaol11.chm3420
f1_keywords:
- vbaol11.chm3420
ms.prod: outlook
api_name:
- Outlook.OlAlwaysDeleteConversation
ms.assetid: 5302003d-b227-5b0b-a8ec-52c107defc97
ms.date: 06/08/2017
localization_priority: Normal
---


# OlAlwaysDeleteConversation enumeration (Outlook)

Specifies constants that determine whether new items of the conversation are always moved to the Deleted Items folder of the specified delivery store.



|Name|Value|Description|
|:-----|:-----|:-----|
| **olAlwaysDelete**|1|New items of the conversation are always moved to the Deleted Items folder for the store that contains the items|
| **olAlwaysDeleteUnsupported**|2|The specified store does not support the action of always moving items to the Deleted Items folder of that store.|
| **olDoNotDelete**|0|New items joining the conversation are not moved to the Deleted Items folder on the specified delivery store, and existing conversation items in the Deleted Items folder are moved to the Inbox.|

## Remarks

This enumeration is used by the [GetAlwaysDelete](Outlook.Conversation.GetAlwaysDelete.md) method of the [Conversation object (Outlook)](Outlook.Conversation.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]