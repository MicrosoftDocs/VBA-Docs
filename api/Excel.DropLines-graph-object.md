---
title: DropLines object (Excel Graph)
keywords: vbagr10.chm5207329
f1_keywords:
- vbagr10.chm5207329
ms.prod: excel
api_name:
- Excel.DropLines
ms.assetid: 52fa64aa-0b0b-bbe1-1ec2-d866e2e35674
ms.date: 04/06/2019
localization_priority: Normal
---


# DropLines object (Excel Graph)

Represents the drop lines in the specified chart group. Drop lines connect the points in the chart with the x-axis. Only line and area chart groups can have drop lines. 

This object isn't a collection. There's no object that represents a single drop line; either you have drop lines turned on for all points in a chart group or you have them turned off.


## Remarks

Use the **[DropLines](excel.droplines-graph-property.md)** property to return the **DropLines** object. 

If the **[HasDropLines](Excel.HasDropLines.md)** property is **False**, most properties of the **DropLines** object are disabled.


## Example

The following example turns on drop lines for chart group one in the chart, and then sets the drop-line color to red.

```vb
myChart.ChartGroups(1).HasDropLines = True 
myChart.ChartGroups(1).DropLines.Border.ColorIndex = 3
```


## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]