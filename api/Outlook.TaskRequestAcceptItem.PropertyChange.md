---
title: TaskRequestAcceptItem.PropertyChange event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.PropertyChange
ms.assetid: 4b26e4b6-607c-c9e6-088f-2e7605b0681f
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestAcceptItem.PropertyChange event (Outlook)

Occurs when an explicit built-in property (for example,  **[Subject](Outlook.AppointmentItem.Subject.md)**) of an instance of the parent object is changed.


## Syntax

_expression_. `PropertyChange`( `_Name_` )

_expression_ A variable that represents a [TaskRequestAcceptItem](Outlook.TaskRequestAcceptItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property that was changed.|

## Remarks

The property name is passed to the event so that you can determine which property was changed.


## See also


[TaskRequestAcceptItem Object](Outlook.TaskRequestAcceptItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]