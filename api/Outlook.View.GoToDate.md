---
title: View.GoToDate method (Outlook)
keywords: vbaol11.chm2496
f1_keywords:
- vbaol11.chm2496
ms.prod: outlook
api_name:
- Outlook.View.GoToDate
ms.assetid: 5ad66fcc-fcdf-9a48-a8e1-669dd294967b
ms.date: 06/08/2017
localization_priority: Normal
---


# View.GoToDate method (Outlook)

Changes the date used by the current view to display information.


## Syntax

_expression_. `GoToDate`( `_Date_` )

_expression_ A variable that represents a [View](Outlook.View.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Date_|Required| **Date**|The date to which the view should be changed.|

## Remarks

To specify a date to go to in a current view, such as a  **[CalendarView](Outlook.calendarView.md)** object, you should first obtain a **[View](Outlook.View.md)** object for the current view by using **[Explorer.CurrentView](Outlook.Explorer.CurrentView.md)** instead of **[Folder.CurrentView](Outlook.Folder.CurrentView.md)**. The following code sample demonstrates how to perform this action.


```vb
Sub TestGoToDate() 
 
 Dim oCV As Outlook.CalendarView 
 
 Dim oExpl As Outlook.Explorer 
 
 Dim datGoTo As Date 
 
 
 
 datGoTo = "11/7/2005" 
 
 
 
 ' Display the contents of the Calendar default folder. 
 
 Set oExpl = Application.Explorers.Add( _ 
 
 Application.Session.GetDefaultFolder(olFolderCalendar), olFolderDisplayFolderOnly) 
 
 oExpl.Display 
 
 
 
 ' Retrieve the current view by calling the 
 
 ' CurrentView property of the Explorer object. 
 
 Set oCV = oExpl.CurrentView 
 
 
 
 ' Set the CalendarViewMode property of the 
 
 ' current view to display items by day. 
 
 oCV.CalendarViewMode = olCalendarViewDay 
 
 
 
 ' Call the GoToDate method to set the date 
 
 ' for which information is displayed in the view. 
 
 oCV.GoToDate datGoTo 
 
End Sub
```


## See also


[View Object](Outlook.View.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]