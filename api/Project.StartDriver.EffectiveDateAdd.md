---
title: StartDriver.EffectiveDateAdd property (Project)
ms.prod: project-server
api_name:
- Project.StartDriver.EffectiveDateAdd
ms.assetid: 5b2e2c6e-06b9-ebf4-efdb-4ca2e944b7ff
ms.date: 06/08/2017
localization_priority: Normal
---


# StartDriver.EffectiveDateAdd property (Project)

Gets the date and time that follows another date by a specified duration, using the effective calendar for a manually scheduled task. Read-only  **Variant**.


## Syntax

_expression_. `EffectiveDateAdd`( `_Date_`, `_Duration_` )

 _expression_ An expression that returns a [StartDriver](./Project.StartDriver.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Date_|Required|**Variant**|Arbitrary date and time, for example, "7/10/2010" or "7/10/2010 2:00:00 PM".|
| _Duration_|Required|**Variant**|Duration to add, for example, "3d" or "2w".|

## Remarks

The  **EffectiveDateAdd** property uses the effective calendar for manually scheduled tasks, which allows tasks to start and finish on non-working times. The property and arguments have no effect on actual task dates.

You can use the  **[EffectiveDateSubtract](Project.StartDriver.EffectiveDateSubtract.md)**, **EffectiveDateAdd**, and **[EffectiveDateDifference](Project.StartDriver.EffectiveDateDifference.md)** properties to calculate start and finish dates for manually scheduled tasks.

To calculate a date for an automatically scheduled task, where you can also specify the calendar, use the  **[DateAdd](Project.Application.DateAdd.md)** method.


## Example

The following statement returns the value "7/9/2009 5:00:00 PM", which is six days after the specified date. 


```vb
Debug.Print ActiveProject.Tasks(3).StartDriver.EffectiveDateAdd("7/2/2009", "6d")
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]