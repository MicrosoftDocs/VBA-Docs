---
title: CalendarView object (Outlook)
keywords: vbaol11.chm3208
f1_keywords:
- vbaol11.chm3208
ms.prod: outlook
api_name:
- Outlook.CalendarView
ms.assetid: 37e078b9-9fc6-5894-b043-06d7257666a8
ms.date: 06/08/2017
localization_priority: Normal
---


# CalendarView object (Outlook)

Represents a view that displays Outlook items in a calendar format.


## Remarks

The **CalendarView** object, derived from the **[View](Outlook.View.md)** object, allows you to create customizable views that allow you to display Outlook items within a calendar, in one of several different modes.

Outlook provides several built-in  **CalendarView** objects, and you can also create custom **CalendarView** objects. Use the **[Add](Outlook.Views.Add.md)** method of the **[Views](Outlook.Views.md)** collection to add a new **CalendarView** to a **[Folder](Outlook.Folder.md)** object. Use the **[Standard](Outlook.TimelineView.Standard.md)** property to determine if an existing **CalendarView** object is built-in or custom.

The **CalendarView** object supports several different view modes, depending on the desired layout and time period in which to display Outlook items. Use the **[CalendarViewMode](Outlook.CalendarView.CalendarViewMode.md)** property to set the view mode, the **[StartField](Outlook.CalendarView.StartField.md)** property to specify the Outlook item property that contains the start date, and the **[EndField](Outlook.CalendarView.EndField.md)** property to specify the Outlook item property that contains the end date for Outlook items to be displayed.

If you set the  **CalendarViewMode** property to any value other than **olCalendarViewMonth**, you can use the **[DayWeekFont](overview/Outlook.md)** and **[DayWeekTimeFont](overview/Outlook.md)** properties to configure the fonts used to display the day, date, and time labels in the view. Use the **[DayWeekTimeScale](Outlook.CalendarView.DayWeekTimeScale.md)** to configure the time scale used to display Outlook items within the view. If you set the **CalendarViewMode** to **olCalendarViewMultiDay**, you can use the **[DaysInMultiDayMode](Outlook.CalendarView.DaysInMultiDayMode.md)** property to determine the number of days to display in the view.

If you set the  **CalendarViewMode** to **olCalendarViewMonth**, you can use the **[MonthFont](overview/Outlook.md)** property to configure the fonts used to display the month and day labels and the **[MonthShowEndTime](Outlook.CalendarView.MonthShowEndTime.md)** to indicate whether the end time for is displayed in the view.

You can also configure how Outlook items appear within the  **CalendarView** object. Use the **[BoldSubjects](Outlook.CalendarView.BoldSubjects.md)** property to indicate whether subjects for Outlook items are displayed in bold and the **[BoldDatesWithItems](Outlook.CalendarView.BoldDatesWithItems.md)** property to indicate whether dates in the Date Navigator that contain Outlook items are displayed in bold. Use the **[Filter](Outlook.CalendarView.Filter.md)** property to determine which Outlook items to display in the view.

The definition for each  **CalendarView** object is stored in Extensible Markup Language (XML) format. Use the **[XML](Outlook.CalendarView.XML.md)** property to work with the XML definition for the **CalendarView** object.

Use the  **[Apply](Outlook.CalendarView.Apply.md)** method to apply any changes made to the **CalendarView** object to the current view. Use the **[Save](Outlook.CalendarView.Save.md)** method to persist any changes made to the **CalendarView** object. Use the **[LockUserChanges](Outlook.CalendarView.LockUserChanges.md)** property to allow or prevent changes to the user interface for the view.

You can change built-in  **CalendarView** objects, but you cannot delete them. Use the **[Delete](Outlook.CalendarView.Delete.md)** method to delete a custom **CalendarView** object. Use the **[Reset](Outlook.CalendarView.Reset.md)** method to reset the properties of a built-in **CalendarView** object to their default values.


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


[Outlook Object Model Reference](overview/Outlook/object-model.md)
[CalendarView Object Members](overview/Outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]