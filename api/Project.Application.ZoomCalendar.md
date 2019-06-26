---
title: Application.ZoomCalendar method (Project)
keywords: vbapj.chm2347
f1_keywords:
- vbapj.chm2347
ms.prod: project-server
api_name:
- Project.Application.ZoomCalendar
ms.assetid: fc02c827-11a0-380b-9e05-b4452246ff05
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ZoomCalendar method (Project)

Zooms in on or out from the Calendar.


## Syntax

_expression_. `ZoomCalendar`( `_NumWeeks_`, `_StartDate_`, `_EndDate_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _NumWeeks_|Optional|**Long**|The number of weeks to display. If StartDate and EndDate are specified, NumWeeks is ignored.|
| _StartDate_|Optional|**Variant**|The first date to display.|
| _EndDate_|Optional|**Variant**|The last date to display.|

## Return value

 **Boolean**


## Remarks

Using the  **ZoomCalendar** method without specifying any arguments displays the **Zoom** dialog box.


## Example

The following example displays four rows of weeks within the active pane of the Calendar view.


```vb
Sub FourWeekCalendar() 
 ViewApply Name:="Calendar" 
 ZoomCalendar NumWeeks:=4 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]