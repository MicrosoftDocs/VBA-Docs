---
title: CalendarView.XML property (Outlook)
keywords: vbaol11.chm2622
f1_keywords:
- vbaol11.chm2622
ms.prod: outlook
api_name:
- Outlook.CalendarView.XML
ms.assetid: f188b827-77c6-71da-0b36-972b16b843a8
ms.date: 06/08/2017
localization_priority: Normal
---


# CalendarView.XML property (Outlook)

Returns or sets a **String** value that specifies the XML definition of the view. Read/write.


## Syntax

_expression_.**XML**

_expression_ A variable that represents a [CalendarView](Outlook.CalendarView.md) object.


## Remarks

The XML definition describes the view type by using a series of tags and keywords corresponding to various properties of the view itself. When the view is created, the XML definition is parsed to render the settings for the new view.

To determine how the XML should be structured when creating views, you can create a view by using the Outlook user interface and then you can retrieve the  **XML** property for that view.


## See also


[CalendarView Object](Outlook.CalendarView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]