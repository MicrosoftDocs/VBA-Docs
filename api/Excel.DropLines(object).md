---
title: DropLines object (Excel)
keywords: vbaxl10.chm603072
f1_keywords:
- vbaxl10.chm603072
ms.prod: excel
api_name:
- Excel.DropLines
ms.assetid: 88fdf5f5-2842-2d68-a073-18d05fd2fa38
ms.date: 03/29/2019
localization_priority: Normal
---


# DropLines object (Excel)

Represents the drop lines in a chart group.


## Remarks

Drop lines connect the points in the chart with the x-axis. Only line and area chart groups can have drop lines. This object isn't a collection. There's no object that represents a single drop line; you either have drop lines turned on for all points in a chart group or you have them turned off.

If the **[HasDropLines](Excel.ChartGroup.HasDropLines.md)** property of the **ChartGroup** object is **False**, most properties of the **DropLines** object are disabled.


## Example

Use the **[DropLines](excel.chartgroup.droplines.md)** property of the **ChartGroup** object to return the **DropLines** object. The following example turns on drop lines for chart group one in embedded chart one, and then sets the drop line color to red.

```vb
Worksheets("sheet1").ChartObjects(1).Activate 
ActiveChart.ChartGroups(1).HasDropLines = True 
ActiveChart.ChartGroups(1).DropLines.Border.ColorIndex = 3
```



## Methods

- [Delete](Excel.DropLines.Delete.md)
- [Select](Excel.DropLines.Select.md)

## Properties

- [Application](Excel.DropLines.Application.md)
- [Border](Excel.DropLines.Border.md)
- [Creator](Excel.DropLines.Creator.md)
- [Format](Excel.DropLines.Format.md)
- [Name](Excel.DropLines.Name.md)
- [Parent](Excel.DropLines.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]