---
title: DocumentItem.PropertyChange event (Outlook)
ms.prod: outlook
api_name:
- Outlook.DocumentItem.PropertyChange
ms.assetid: ec757f98-db44-585e-1a4a-5b3044428dec
ms.date: 06/08/2017
localization_priority: Normal
---


# DocumentItem.PropertyChange event (Outlook)

Occurs when an explicit built-in property (for example,  **[Subject](Outlook.AppointmentItem.Subject.md)**) of an instance of the parent object is changed.


## Syntax

_expression_. `PropertyChange`( `_Name_` )

_expression_ A variable that represents a [DocumentItem](Outlook.DocumentItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the property that was changed.|

## Remarks

The property name is passed to the event so that you can determine which property was changed.


## See also


[DocumentItem Object](Outlook.DocumentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]