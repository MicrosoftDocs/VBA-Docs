---
title: Application.CalendarBarStyles method (Project)
keywords: vbapj.chm2326
f1_keywords:
- vbapj.chm2326
ms.prod: project-server
api_name:
- Project.Application.CalendarBarStyles
ms.assetid: bf168abd-3033-f187-ee3e-19e672be4aac
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.CalendarBarStyles method (Project)

Turns bar rounding on or off in the Calendar.


## Syntax

_expression_. `CalendarBarStyles`( `_BarRounding_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _BarRounding_|Optional|**Boolean**|**True** if bars in the Calendar round to midnight if their start times are less than or equal to the default start time, or if their end times are greater than or equal to the default end time. If **BarRounding** is omitted, the **Bar Styles** dialog box appears.|

## Return value

 **Boolean**


## Remarks

The default start and default end times can be set with the  **OptionsCalendar** method.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]