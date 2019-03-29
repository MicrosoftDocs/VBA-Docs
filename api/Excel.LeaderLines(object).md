---
title: LeaderLines object (Excel)
keywords: vbaxl10.chm605072
f1_keywords:
- vbaxl10.chm605072
ms.prod: excel
api_name:
- Excel.LeaderLines
ms.assetid: ff4954f1-6967-9dd8-c9d6-8927d079e995
ms.date: 03/30/2019
localization_priority: Normal
---


# LeaderLines object (Excel)

Represents leader lines on a chart. Leader lines connect data labels to data points.


## Remarks

This object isn't a collection; there's no object that represents a single leader line.

This object applies only to pie charts.


## Example

Use the **[LeaderLines](Excel.Series.LeaderLines.md)** property of the **Series** object to return the **LeaderLines** object. 

The following example adds data labels and blue leader lines to series one on chart one. If no leader lines are visible, this example code will fail. In this situation, you can manually drag one of the data labels away from the pie chart to make a leader line show up.

```vb
With Worksheets(1).ChartObjects(1).Chart.SeriesCollection(1) 
 .HasDataLabels = True 
 .DataLabels.Position = xlLabelPositionBestFit 
 .HasLeaderLines = True 
 .LeaderLines.Border.ColorIndex = 5 
End With
```

## Methods

- [Delete](Excel.LeaderLines.Delete.md)
- [Select](Excel.LeaderLines.Select.md)

## Properties

- [Application](Excel.LeaderLines.Application.md)
- [Border](Excel.LeaderLines.Border.md)
- [Creator](Excel.LeaderLines.Creator.md)
- [Format](Excel.LeaderLines.Format.md)
- [Parent](Excel.LeaderLines.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]