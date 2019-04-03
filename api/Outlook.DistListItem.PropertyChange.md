---
title: DistListItem.PropertyChange event (Outlook)
ms.prod: outlook
api_name:
- Outlook.DistListItem.PropertyChange
ms.assetid: 932a2ded-1e92-e40d-8f88-9044cccb7863
ms.date: 06/08/2017
localization_priority: Normal
---


# DistListItem.PropertyChange event (Outlook)

Occurs when an explicit built-in property (for example,  **[Subject](Outlook.AppointmentItem.Subject.md)**) of an instance of the parent object is changed.


## Syntax

_expression_. `PropertyChange`( `_Name_` )

_expression_ A variable that represents a [DistListItem](Outlook.DistListItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property that was changed.|

## Remarks

The property name is passed to the event so that you can determine which property was changed.


## See also


[DistListItem Object](Outlook.DistListItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]