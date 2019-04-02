---
title: MeetingItem.BeforeAutoSave event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MeetingItem.BeforeAutoSave
ms.assetid: 59de272e-a36a-e842-a962-03ebe2befa26
ms.date: 06/08/2017
localization_priority: Normal
---


# MeetingItem.BeforeAutoSave event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

_expression_. `BeforeAutoSave`( `_Cancel_` )

_expression_ A variable that represents a [MeetingItem](Outlook.MeetingItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[MeetingItem](Outlook.MeetingItem.md)** to be saved.|

## See also


[MeetingItem Object](Outlook.MeetingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]