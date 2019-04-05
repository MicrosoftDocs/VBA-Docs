---
title: DataLabels collection (Excel Graph)
keywords: vbagr10.chm131187
f1_keywords:
- vbagr10.chm131187
ms.prod: excel
ms.assetid: 597c7269-71ed-5dcc-af6b-34dc908e9d58
ms.date: 04/06/2019
localization_priority: Normal
---


# DataLabels collection (Excel Graph)

A collection of all the **[DataLabel](Excel.DataLabel-graph-object.md)** objects for the specified series. Each **DataLabel** object represents a data label for a point or trendline. For a series without definable points (such as an area series), the **DataLabels** collection contains a single data label.


## Remarks

Use the **[DataLabels](excel.datalabels-graph-method.md)** method to return the **DataLabels** collection. 

Use **DataLabels** (_index_), where _index_ is the data label's index number, to return a single **DataLabel** object.

## Example

The following example sets the number format for data labels in series one in the chart.

```vb
With myChart.SeriesCollection(1) 
 .HasDataLabels = True 
 .DataLabels.NumberFormat = "##.##" 
End With
```

<br/>

The following example sets the number format for the fifth data label in series one in the chart.

```vb
myChart.SeriesCollection(1).DataLabels(5).NumberFormat = "0.000"
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]