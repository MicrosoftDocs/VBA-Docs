---
title: TaskRequestUpdateItem.PropertyChange event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem.PropertyChange
ms.assetid: 47121ba2-cd73-405a-9bd0-d8fc4a77a535
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestUpdateItem.PropertyChange event (Outlook)

Occurs when an explicit built-in property (for example,  **[Subject](Outlook.AppointmentItem.Subject.md)**) of an instance of the parent object is changed.


## Syntax

_expression_. `PropertyChange`( `_Name_` )

_expression_ A variable that represents a [TaskRequestUpdateItem](Outlook.TaskRequestUpdateItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property that was changed.|

## Remarks

The property name is passed to the event so that you can determine which property was changed.


## See also


[TaskRequestUpdateItem Object](Outlook.TaskRequestUpdateItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]