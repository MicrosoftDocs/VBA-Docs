---
title: MeetingItem.BeforeDelete event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MeetingItem.BeforeDelete
ms.assetid: 155c5225-aeb0-55b6-26dc-811d00128238
ms.date: 06/08/2017
localization_priority: Normal
---


# MeetingItem.BeforeDelete event (Outlook)

Occurs before an item (which is an instance of the parent object) is deleted.


## Syntax

_expression_.**BeforeDelete** (_Item_, _Cancel_)

_expression_ A variable that represents a [MeetingItem](Outlook.MeetingItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|The item being deleted.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the operation is not completed and the item is not deleted.|

## Remarks

In order for this event to fire when an email message, distribution list, journal entry, task, contact, or post are deleted through an action, an inspector must be open.

The event occurs each time an item is deleted.


## See also


[MeetingItem Object](Outlook.MeetingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]