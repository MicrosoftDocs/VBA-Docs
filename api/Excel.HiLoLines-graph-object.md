---
title: HiLoLines object (Excel Graph)
keywords: vbagr10.chm5207532
f1_keywords:
- vbagr10.chm5207532
ms.prod: excel
api_name:
- Excel.HiLoLines
ms.assetid: 6793025e-0b3e-360c-4292-02397395535a
ms.date: 04/06/2019
localization_priority: Normal
---


# HiLoLines object (Excel Graph)

Represents the high-low lines in the specified chart group. High-low lines connect the highest point with the lowest point in every category in the chart group. Only 2D line groups can have high-low lines. 

This object isn't a collection. There's no object that represents a single high-low line; either you have high-low lines turned on for all points in a chart group or you have them turned off.


## Remarks

Use the **[HiLoLines](excel.hilolines-graph-property.md)** property to return the **HiLoLines** object. 

If the **[HasHiLoLines](Excel.HasHiLoLines.md)** property is **False**, most properties of the **HiLoLines** object are disabled.


## Example

The following example makes the high-low lines in chart group one in the chart blue.

```vb
myChart.ChartGroups(1).HiLoLines.Border.Color = RGB(0, 0, 255)
```


## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]