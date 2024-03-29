---
title: Gridlines object (Excel Graph)
keywords: vbagr10.chm131203
f1_keywords:
- vbagr10.chm131203
api_name:
- Excel.Gridlines
ms.assetid: 8879cdea-609f-5994-3fb6-3a9d5fa849b4
ms.date: 04/06/2019
ms.localizationpriority: medium
---


# Gridlines object (Excel Graph)

Represents major or minor gridlines on the specified chart axis. Gridlines extend the tick marks on a chart axis to make it easier to see the values associated with the data markers. 

This object isn't a collection. There's no object that represents a single gridline; either you have all gridlines for an axis turned on or you have them all turned off.


## Remarks

Use the **[MajorGridlines](excel.majorgridlines.md)** property to return the **GridLines** object that represents the major gridlines for the axis. 

Use the **[MinorGridlines](excel.minorgridlines.md)** property to return the **GridLines** object that represents the minor gridlines for the axis. It's possible to return both major and minor gridlines at the same time.

## Example

The following example turns on major gridlines for the category axis on the chart, and then formats the gridlines to be blue dashed lines.

```vb
With myChart.Axes(xlCategory) 
 .HasMajorGridlines = True 
 .MajorGridlines.Border.Color = RGB(0, 0, 255) 
 .MajorGridlines.Border.LineStyle = xlDash 
End With
```

## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]