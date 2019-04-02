---
title: RemoteItem.PropertyChange event (Outlook)
ms.prod: outlook
api_name:
- Outlook.RemoteItem.PropertyChange
ms.assetid: 630d4423-cb56-eef0-e1b1-1afe227c140d
ms.date: 06/08/2017
localization_priority: Normal
---


# RemoteItem.PropertyChange event (Outlook)

Occurs when an explicit built-in property (for example,  **[Subject](Outlook.AppointmentItem.Subject.md)**) of an instance of the parent object is changed.


## Syntax

_expression_. `PropertyChange`( `_Name_` )

_expression_ A variable that represents a [RemoteItem](Outlook.RemoteItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property that was changed.|

## Remarks

The property name is passed to the event so that you can determine which property was changed.


## See also


[RemoteItem Object](Outlook.RemoteItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]