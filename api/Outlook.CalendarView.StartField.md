---
title: CalendarView.StartField property (Outlook)
keywords: vbaol11.chm2625
f1_keywords:
- vbaol11.chm2625
ms.prod: outlook
api_name:
- Outlook.CalendarView.StartField
ms.assetid: 085c6605-0bff-98a5-fb48-ce32b76037db
ms.date: 06/08/2017
localization_priority: Normal
---


# CalendarView.StartField property (Outlook)

Returns or sets a  **String** value that represents the name of the property that starts the time duration for Outlook items displayed in the **[CalendarView](Outlook.CalendarView.md)** object. Read/write.


## Syntax

_expression_. `StartField`

_expression_ A variable that represents a [CalendarView](Outlook.CalendarView.md) object.


## Remarks

The values of the  **StartField** and **[EndField](Outlook.CalendarView.EndField.md)** properties indicate which Outlook item properties are used by the **CalendarView** object to represent the duration of an Outlook item. Both custom and built-in properties can be specified, but only date/time properties are allowed.


## See also


[CalendarView Object](Outlook.CalendarView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]