---
title: Application.DateAdd method (Project)
ms.prod: project-server
api_name:
- Project.Application.DateAdd
ms.assetid: df0da054-495c-c224-ebc8-b47acb78e2af
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DateAdd method (Project)

Returns the date and time that follows another date by a specified duration, for an automatically scheduled task.


## Syntax

_expression_. `DateAdd`( `_StartDate_`, `_Duration_`, `_Calendar_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _StartDate_|Required|**Variant**|The original date to which the duration is added.|
| _Duration_|Required|**Variant**|The duration to add to the start date.|
| _Calendar_|Optional|**Object**|A resource, task, or base calendar object. The default value is the calendar of the active project.|

## Return value

 **Variant**


## Remarks

To add a duration to a date for a manually scheduled task, which uses an effective calendar that can include non-working time, use the **[EffectiveDateAdd](Project.StartDriver.EffectiveDateAdd.md)** property.


## Example

The following example displays the finish date of a three-day automatically scheduled task that begins on 7/11/07 at 8 A.M.


```vb
Sub FindFinishDate() 
 MsgBox Application.DateAdd(StartDate:="7/11/07 8:00 AM", Duration:="3d") 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]