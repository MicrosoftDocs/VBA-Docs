---
title: CalendarView.EndField property (Outlook)
keywords: vbaol11.chm2626
f1_keywords:
- vbaol11.chm2626
ms.prod: outlook
api_name:
- Outlook.CalendarView.EndField
ms.assetid: 311994db-ef43-e49c-6f0e-9b346d0bb3ca
ms.date: 06/08/2017
localization_priority: Normal
---


# CalendarView.EndField property (Outlook)

Returns or sets a  **String** value that represents the name of the property that ends the time duration for Outlook items displayed in the **[CalendarView](Outlook.CalendarView.md)** object. Read/write.


## Syntax

_expression_. `EndField`

_expression_ A variable that represents a [CalendarView](Outlook.CalendarView.md) object.


## Remarks

The values of the  **[StartField](Outlook.CalendarView.StartField.md)** and **EndField** properties indicate which Outlook item properties are used by the **CalendarView** object to represent the duration of an Outlook item. Both custom and built-in properties can be specified, but only date/time properties are allowed.


## See also


[CalendarView Object](Outlook.CalendarView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]