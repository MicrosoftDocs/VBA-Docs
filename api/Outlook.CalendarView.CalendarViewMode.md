---
title: CalendarView.CalendarViewMode property (Outlook)
keywords: vbaol11.chm2627
f1_keywords:
- vbaol11.chm2627
ms.prod: outlook
api_name:
- Outlook.CalendarView.CalendarViewMode
ms.assetid: 144e46ed-984f-fac0-fad3-0ff5ac9f2996
ms.date: 06/08/2017
localization_priority: Normal
---


# CalendarView.CalendarViewMode property (Outlook)

Returns or sets an  **[OlCalendarViewMode](Outlook.OlCalendarViewMode.md)** that determines the current view mode of the **[CalendarView](Outlook.CalendarView.md)** object. Read/write.


## Syntax

_expression_. `CalendarViewMode`

_expression_ A variable that represents a [CalendarView](Outlook.CalendarView.md) object.


## Example

The following Visual Basic for Applications (VBA) example configures the current  **CalendarView** object to show a single day, using an 8-point Verdana font to display items and a 16-point Verdana font to display time values and the Tasks header within the view.


```vb
Sub ConfigureDayViewFonts() 
 
 Dim objView As CalendarView 
 
 
 
 ' Check if the current view is a calendar view. 
 
 If Application.ActiveExplorer.CurrentView.ViewType = _ 
 
 olCalendarView Then 
 
 
 
 ' Obtain a CalendarView object reference for the 
 
 ' current calendar view. 
 
 Set objView = _ 
 
 Application.ActiveExplorer.CurrentView 
 
 
 
 With objView 
 
 ' Set the calendar view to show a 
 
 ' single day. 
 
 .CalendarViewMode = olCalendarViewDay 
 
 
 
 ' Set the DayWeekFont to 8-point Verdana. 
 
 .DayWeekFont.Name = "Verdana" 
 
 .DayWeekFont.Size = 8 
 
 
 
 ' Set the DayWeekTimeFont to 16-point Verdana. 
 
 .DayWeekTimeFont.Name = "Verdana" 
 
 .DayWeekTimeFont.Size = 16 
 
 
 
 ' Save the calendar view. 
 
 .Save 
 
 End With 
 
 End If 
 
End Sub
```


## See also


[CalendarView Object](Outlook.CalendarView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]