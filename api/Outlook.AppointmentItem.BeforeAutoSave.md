---
title: AppointmentItem.BeforeAutoSave event (Outlook)
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.BeforeAutoSave
ms.assetid: c24e39d1-39e5-6422-78ff-9d4e391ea2ae
ms.date: 06/08/2017
localization_priority: Normal
---


# AppointmentItem.BeforeAutoSave event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

_expression_. `BeforeAutoSave`( `_Cancel_` , )

_expression_ A variable that represents an [AppointmentItem](Outlook.AppointmentItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[AppointmentItem](Outlook.AppointmentItem.md)** to be saved.|

## See also


[AppointmentItem Object](Outlook.AppointmentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]