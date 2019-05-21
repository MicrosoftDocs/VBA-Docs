---
title: RemoteItem.BeforeDelete event (Outlook)
ms.prod: outlook
api_name:
- Outlook.RemoteItem.BeforeDelete
ms.assetid: 0f1f4b6d-7a5a-2302-2b71-eea7bf7f1af9
ms.date: 06/08/2017
localization_priority: Normal
---


# RemoteItem.BeforeDelete event (Outlook)

Occurs before an item (which is an instance of the parent object) is deleted.


## Syntax

_expression_.**BeforeDelete** (_Item_, _Cancel_)

_expression_ A variable that represents a [RemoteItem](Outlook.RemoteItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|The item being deleted.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the operation is not completed and the item is not deleted.|

## Remarks

In order for this event to fire when an email message, distribution list, journal entry, task, contact, or post are deleted through an action, an inspector must be open.

The event occurs each time an item is deleted.


## See also


[RemoteItem Object](Outlook.RemoteItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]