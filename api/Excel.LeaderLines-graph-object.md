---
title: LeaderLines object (Excel Graph)
keywords: vbagr10.chm5207590
f1_keywords:
- vbagr10.chm5207590
ms.prod: excel
api_name:
- Excel.LeaderLines
ms.assetid: 9704f195-dbbc-6979-c57d-8ced3557cdde
ms.date: 04/06/2019
localization_priority: Normal
---


# LeaderLines object (Excel Graph)

Represents leader lines in the specified chart. Leader lines connect data labels to data points. 

This object isn't a collection; there's no object that represents a single leader line.


## Remarks

Use the **[LeaderLines](Excel.LeaderLines-graph-property.md)** property to return the **LeaderLines** object. 



## Example

The following example adds data labels and blue leader lines to series one in the chart.

```vb
With myChart.SeriesCollection(1) 
 .HasDataLabels = True 
 .DataLabels.Position = xlLabelPositionBestFit 
 .HasLeaderLines = True 
 .LeaderLines.Border.ColorIndex = 5 
End With
```

## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]