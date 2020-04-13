---
title: Application.PanZoomZoomTo method (Project)
ms.prod: project-server
api_name:
- Project.Application.PanZoomZoomTo
ms.assetid: bd8510b8-fbdb-2c96-94a7-98c377b2d331
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.PanZoomZoomTo method (Project)

Zooms the time range in or out for the Gantt chart in the active view.


## Syntax

_expression_. `PanZoomZoomTo`( `_Start_`, `_Finish_` )

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Start_|Required|**Variant**|Specifies the start date for the left side of the Gantt chart.|
| _Finish_|Required|**Variant**|Specifies the finish date for the right side of the Gantt chart.|

## Return value

Nothing


## Remarks

The **PanZoomZoomTo** method has no effect on the Calendar view or the Network Diagram (PERT chart) view.

To pan the Gantt chart to a different starting date and maintain the same timescale, use the **[PanZoomPanTo](Project.Application.PanZoomPanTo.md)** method. To change the timescale format and labels, use the **[TimescaleEdit](Project.Application.TimescaleEdit.md)** method.


## Example

The following command zooms in the Gantt chart to show the time between 10 AM and 5 PM on March 19, 2012.


```vb
PanZoomZoomTo Start:="3/19/2012 10:00:00 AM", Finish:="3/19/2012 5:00:00 PM"
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]