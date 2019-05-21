---
title: AppointmentItem.BeforeDelete event (Outlook)
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.BeforeDelete
ms.assetid: dc6944f6-e020-bdd7-0b64-98a3f3d2e94c
ms.date: 06/08/2017
localization_priority: Normal
---


# AppointmentItem.BeforeDelete event (Outlook)

Occurs before an item (which is an instance of the parent object) is deleted.


## Syntax

_expression_.**BeforeDelete** (_Item_, _Cancel_)

_expression_ A variable that represents an [AppointmentItem](Outlook.AppointmentItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|The item being deleted.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the operation is not completed and the item is not deleted.|

## Remarks

In order for this event to fire when an email message, distribution list, journal entry, task, contact, or post are deleted through an action, an inspector must be open.

The event occurs each time an item is deleted.


## See also


[AppointmentItem Object](Outlook.AppointmentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]