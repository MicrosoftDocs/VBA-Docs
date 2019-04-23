---
title: SharingItem.PropertyChange event (Outlook)
ms.prod: outlook
api_name:
- Outlook.SharingItem.PropertyChange
ms.assetid: 7c3cf73a-4b2c-3f74-4d3e-5a0e04870f07
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.PropertyChange event (Outlook)

Occurs when an explicit built-in property (for example,  **[Subject](Outlook.SharingItem.Subject.md)**) of an instance of the parent object is changed.


## Syntax

_expression_. `PropertyChange`( `_Name_` )

 _expression_ An expression that returns a [SharingItem](Outlook.SharingItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property that was changed.|

## Remarks

The property name is passed to the event so that you can determine which property was changed.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]