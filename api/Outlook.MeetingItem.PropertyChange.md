---
title: MeetingItem.PropertyChange event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MeetingItem.PropertyChange
ms.assetid: 6bc3629b-b08a-0d8b-f1e3-6d3c90176ac2
ms.date: 06/08/2017
localization_priority: Normal
---


# MeetingItem.PropertyChange event (Outlook)

Occurs when an explicit built-in property (for example,  **[Subject](Outlook.AppointmentItem.Subject.md)**) of an instance of the parent object is changed.


## Syntax

_expression_. `PropertyChange`( `_Name_` )

_expression_ A variable that represents a [MeetingItem](Outlook.MeetingItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property that was changed.|

## Remarks

The property name is passed to the event so that you can determine which property was changed.


## See also


[MeetingItem Object](Outlook.MeetingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]