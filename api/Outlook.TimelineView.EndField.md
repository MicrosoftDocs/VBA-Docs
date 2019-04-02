---
title: TimelineView.EndField property (Outlook)
keywords: vbaol11.chm2661
f1_keywords:
- vbaol11.chm2661
ms.prod: outlook
api_name:
- Outlook.TimelineView.EndField
ms.assetid: 7fef24ee-f96a-39e5-5b9a-9fe46ee7c627
ms.date: 06/08/2017
localization_priority: Normal
---


# TimelineView.EndField property (Outlook)

Returns or sets a  **String** value that represents the name of the property that ends the time duration for Outlook items displayed in the **[TimelineView](Outlook.TimelineView.md)** object. Read/write.


## Syntax

_expression_. `EndField`

_expression_ A variable that represents a [TimelineView](Outlook.TimelineView.md) object.


## Remarks

The values of the  **[StartField](Outlook.TimelineView.StartField.md)** and **EndField** properties indicate which Outlook item properties are used by the **TimelineView** object to represent the duration of an Outlook item. Both custom and built-in properties can be specified, but only date/time properties are allowed.


## See also


[TimelineView Object](Outlook.TimelineView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]