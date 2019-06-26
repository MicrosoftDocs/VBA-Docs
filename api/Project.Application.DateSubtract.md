---
title: Application.DateSubtract method (Project)
ms.prod: project-server
api_name:
- Project.Application.DateSubtract
ms.assetid: 1eb05a59-271d-31d0-8945-23bc3c9600e0
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DateSubtract method (Project)

Returns the date and time that precedes another date by a specified duration, for an automatically scheduled task.


## Syntax

_expression_. `DateSubtract`( `_FinishDate_`, `_Duration_`, `_Calendar_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FinishDate_|Required|**Variant**|The date used as the end of the duration.|
| _Duration_|Required|**Variant**|The duration to subtract from the finish date.|
| _Calendar_|Optional|**Object**|A resource, task, or base calendar object. The default value is the calendar of the active project.|

## Return value

 **Variant**


## Remarks

To subtract a duration from a date for a manually scheduled task, which uses an effective calendar that can include non-working time, use the  **[EffectiveDateSubtract](Project.StartDriver.EffectiveDateSubtract.md)** property.


## Example

The following example displays the start date of a task that lasts three days and ends on 7/13/02 at 5:00 P.M.


```vb
Sub FindDuration() 
 MsgBox DateSubtract("7/13/02 5:00 PM", "3d") 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]