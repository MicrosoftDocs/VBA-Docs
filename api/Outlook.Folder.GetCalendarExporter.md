---
title: Folder.GetCalendarExporter method (Outlook)
keywords: vbaol11.chm2020
f1_keywords:
- vbaol11.chm2020
ms.prod: outlook
api_name:
- Outlook.Folder.GetCalendarExporter
ms.assetid: 7c67e208-65dd-8904-4b6f-8ec2df4e530d
ms.date: 06/08/2017
localization_priority: Normal
---


# Folder.GetCalendarExporter method (Outlook)

Creates a  **[CalendarSharing](Outlook.CalendarSharing.md)** object for the specified **[Folder](Outlook.Folder.md)**.


## Syntax

_expression_. `GetCalendarExporter`

 _expression_ An expression that returns a [Folder](Outlook.Folder.md) object.


## Return value

A  **CalendarSharing** object for the specified folder.


## Remarks

The  **GetCalendarExporter** method automatically sets the defaults for the **CalendarSharing** class to the standard default options used by the **Folder** object. The **GetCalendarExporter** method can only be used on calendar folders. An error occurs if you use the method on **Folder** objects that represent other folder types.


> [!NOTE] 
> The  **CalendarSharing** object only supports exporting the iCalendar (.ics) file format.


## See also


[Folder Object](Outlook.Folder.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]