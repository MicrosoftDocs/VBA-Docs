---
title: SeriesCollection collection (Excel Graph)
keywords: vbagr10.chm131116
f1_keywords:
- vbagr10.chm131116
ms.prod: excel
ms.assetid: c5d00466-f7a1-7e6f-56e4-958901dbe3e3
ms.date: 04/06/2019
localization_priority: Normal
---


# SeriesCollection collection (Excel Graph)

A collection of all the **[Series](Excel.Series-graph-object.md)** objects in the specified chart or chart group.


## Remarks

Use the **[SeriesCollection](excel.seriescollection-graph-method.md)** method to return the **SeriesCollection** collection. 

Use **SeriesCollection** (_index_), where _index_ is the series' index number or name, to return a single **Series** object. 

## Example

The following example adjusts the interior color for each series in the collection.

```vb
For X = 1 To myChart.SeriesCollection.Count 
 With myChart.SeriesCollection(X) 
 .Interior.Color = RGB(X * 75, 50, X * 50) 
 End With 
Next X
```

<br/>

The following example sets the color of the interior for series one in the chart to red.

```vb
myChart.SeriesCollection(1).Interior.Color = RGB(255, 0, 0)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]