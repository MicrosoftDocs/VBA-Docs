---
title: UpBars object (Excel)
keywords: vbaxl10.chm607072
f1_keywords:
- vbaxl10.chm607072
ms.prod: excel
api_name:
- Excel.UpBars
ms.assetid: 4f2a85fe-3fbb-ccc6-7b16-e48e54cd3394
ms.date: 04/03/2019
localization_priority: Normal
---


# UpBars object (Excel)

Represents the up bars in a chart group.


## Remarks

Up bars connect points on series one with higher values on the last series in the chart group (the lines go up from series one). Only 2D line groups that contain at least two series can have up bars. This object isn't a collection. There's no object that represents a single up bar; you either have up bars turned on for all points in a chart group or you have them turned off.

If the **[HasUpDownBars](Excel.ChartGroup.HasUpDownBars.md)** property of the **ChartGroup** object is **False**, most properties of the **UpBars** object are disabled.


## Example

Use the **[UpBars](Excel.ChartGroup.UpBars.md)** property of the **ChartGroup** object to return the **UpBars** object.

The following example turns on up and down bars for chart group one in embedded chart one on Sheet5. The example then sets the up bar color to blue and sets the down bar color to red.

```vb
With Worksheets("sheet5").ChartObjects(1).Chart.ChartGroups(1) 
 .HasUpDownBars = True 
 .UpBars.Interior.Color = RGB(0, 0, 255) 
 .DownBars.Interior.Color = RGB(255, 0, 0) 
End With
```

## Methods

- [Delete](Excel.UpBars.Delete.md)
- [Select](Excel.UpBars.Select.md)

## Properties

- [Application](Excel.UpBars.Application.md)
- [Creator](Excel.UpBars.Creator.md)
- [Format](Excel.UpBars.Format.md)
- [Name](Excel.UpBars.Name.md)
- [Parent](Excel.UpBars.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]