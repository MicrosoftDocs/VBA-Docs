---
title: Application.CalendarBestFitWeekHeight method (Project)
keywords: vbapj.chm2327
f1_keywords:
- vbapj.chm2327
ms.prod: project-server
api_name:
- Project.Application.CalendarBestFitWeekHeight
ms.assetid: 58b7e8e8-4001-ef47-c7ba-71af617768eb
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.CalendarBestFitWeekHeight method (Project)

Changes the height of the active calendar box to display all task bars.


## Syntax

_expression_. `CalendarBestFitWeekHeight`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Return value

 **Boolean**


## Example

The following example changes the height of the active calendar box to display all task bars. 


```vb
Sub CalendarBestFit_WeekHeight() 
 
 Dim Result As Boolean 
 
 'Activate Caldender view 
 ViewApply Name:="Calendar" 
 Result = CalendarBestFitWeekHeight() 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]