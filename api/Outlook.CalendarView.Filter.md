---
title: CalendarView.Filter property (Outlook)
keywords: vbaol11.chm2624
f1_keywords:
- vbaol11.chm2624
ms.prod: outlook
api_name:
- Outlook.CalendarView.Filter
ms.assetid: c62e9521-e1aa-bfe8-5774-25c3227973b5
ms.date: 06/08/2017
localization_priority: Normal
---


# CalendarView.Filter property (Outlook)

Returns or sets a **String** value that represents the filter for a view. Read/write.


## Syntax

_expression_. `Filter`

_expression_ A variable that represents a [CalendarView](Outlook.CalendarView.md) object.


## Remarks

The value of this property is a string, in DAV Searching and Locating (DASL) syntax, that represents the current filter for the view. For more information about using DASL syntax to filter items in a view, see [Filtering Items](../outlook/How-to/Search-and-Filter/filtering-items.md).


## Example

The following Visual Basic for Applications (VBA) example obtains a **[View](Outlook.View.md)** object by using the **[CurrentView](Outlook.Explorer.CurrentView.md)** property of the **[Explorer](Outlook.Explorer.md)** object, then sets the **[Filter](Outlook.View.Filter.md)** property of the **View** object to display only those Outlook items that were received last week.


```vb
Private Sub FilterViewToLastWeek() 
 
 Dim objView As View 
 
 
 
 ' Obtain a View object reference to the current view. 
 
 Set objView = Application.ActiveExplorer.CurrentView 
 
 
 
 ' Set a DASL filter string, using a DASL macro, to show 
 
 ' only those items that were received last week. 
 
 objView.Filter = "%lastweek(""urn:schemas:httpmail:datereceived"")%" 
 
 
 
 ' Save and apply the view. 
 
 objView.Save 
 
 objView.Apply 
 
End Sub
```


## See also


[CalendarView Object](Outlook.CalendarView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]