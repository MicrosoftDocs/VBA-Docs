---
title: Border object (Graph)
keywords: vbagr10.chm5207143
f1_keywords:
- vbagr10.chm5207143
ms.prod: excel
api_name:
- Excel.Border
ms.assetid: cb5ee6ef-f497-5113-85e4-a312871ad072
ms.date: 03/07/2019
localization_priority: Normal
---


# Border object (Graph)

Represents the border of the specified object.

## Remarks

An object's border is treated as a single entity and is always returned as a unit (in its entirety), regardless of how many sides it has. 

Use the **[Border](Excel.Border-graph-property.md)** property to return the **Border** object. 

## Example

The following example places a dashed border around the chart area and places a dotted border around the plot area.

```vb
With myChart 
 .ChartArea.Border.LineStyle = xlDash 
 .PlotArea.Border.LineStyle = xlDot 
End With
```

## See also

- [Border object (Excel)](Excel.Border(object).md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]