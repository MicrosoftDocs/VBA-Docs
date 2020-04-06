---
title: Calendar.Period method (Project)
ms.prod: project-server
api_name:
- Project.Calendar.Period
ms.assetid: b717bcbe-654b-5791-2002-d65e2a96617f
ms.date: 06/08/2017
localization_priority: Normal
---


# Calendar.Period method (Project)

Gets a **[Period](Project.Period.md)** object representing a period of time in a calendar. Read-only **Period**.


## Syntax

_expression_. `Period`( `_Start_`, `_Finish_` )

_expression_ A variable that represents a [Calendar](./Project.Calendar.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Start_|Required|**Variant**|The start date of the desired period.|
| _Finish_|Optional|**Variant**| The finish date of the desired period. The default value is the same date as Start.|

## Return value

 **Period**


## Example

The following example sets a winter holiday for the active project.


```vb
Sub SetWinterHoliday() 
    ActiveProject.Calendar.Period("12/20/02", "12/31/02").Working = False 
 End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]