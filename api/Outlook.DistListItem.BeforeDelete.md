---
title: DistListItem.BeforeDelete event (Outlook)
ms.prod: outlook
api_name:
- Outlook.DistListItem.BeforeDelete
ms.assetid: 095be9b1-b7c5-bc22-a478-4f4dfeb91641
ms.date: 06/08/2017
localization_priority: Normal
---


# DistListItem.BeforeDelete event (Outlook)

Occurs before an item (which is an instance of the parent object) is deleted.


## Syntax

_expression_.**BeforeDelete** (_Item_, _Cancel_)

_expression_ A variable that represents a [DistListItem](Outlook.DistListItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|The item being deleted.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the operation is not completed and the item is not deleted.|

## Remarks

In order for this event to fire when an email message, distribution list, journal entry, task, contact, or post are deleted through an action, an inspector must be open.

The event occurs each time an item is deleted.


## See also


[DistListItem Object](Outlook.DistListItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]