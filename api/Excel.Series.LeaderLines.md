---
title: Series.LeaderLines property (Excel)
keywords: vbaxl10.chm578121
f1_keywords:
- vbaxl10.chm578121
api_name:
- Excel.Series.LeaderLines
ms.assetid: d08a982c-8ac0-3f72-3f94-d72b3081f013
ms.date: 05/11/2019
ms.localizationpriority: medium
---


# Series.LeaderLines property (Excel)

Returns a **[LeaderLines](excel.leaderlines(object).md)** object that represents the leader lines for the series. Read-only.


## Syntax

_expression_.**LeaderLines**

_expression_ A variable that represents a **[Series](Excel.Series(object).md)** object.


## Remarks

This property applies only to pie charts.


## Example

This example adds data labels and blue leader lines to series one on the pie chart. If no leader lines are visible, this example code will fail. In this situation, you can manually drag one of the data labels away from the pie chart to make a leader line show up.

```vb
With Worksheets(1).ChartObjects(1).Chart.SeriesCollection(1) 
 .HasDataLabels = True 
 .DataLabels.Position = xlLabelPositionBestFit 
 .HasLeaderLines = True 
 .LeaderLines.Border.ColorIndex = 5 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]