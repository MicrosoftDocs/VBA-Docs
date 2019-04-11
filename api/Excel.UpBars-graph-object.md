---
title: UpBars object (Excel Graph)
keywords: vbagr10.chm5208101
f1_keywords:
- vbagr10.chm5208101
ms.prod: excel
api_name:
- Excel.UpBars
ms.assetid: 635f449d-eb8b-2026-e1a7-9472f33641cc
ms.date: 04/06/2019
localization_priority: Normal
---


# UpBars object (Excel Graph)

Represents the up bars in a chart group. Up bars connect points in series one with higher values in the last series in the chart group (the lines go up from series one). Only 2D line groups that contain at least two series can have up bars. 

This object isn't a collection. There's no object that represents a single up bar; either you have up bars turned on for all points in a chart group or you have them turned off.


## Remarks

Use the **[UpBars](excel.upbars-graph-property.md)** property to return the **UpBars** object. 

If the **[HasUpDownBars](Excel.HasUpDownBars.md)** property is **False**, most properties of the **UpBars** object are disabled.

## Example

The following example turns on up and down bars for chart group one in the chart. The example then sets the up-bar color to blue and sets the down-bar color to red.

```vb
With myChart.ChartGroups(1) 
 .HasUpDownBars = True 
 .UpBars.Interior.Color = RGB(0, 0, 255) 
 .DownBars.Interior.Color = RGB(255, 0, 0) 
End With
```


## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]