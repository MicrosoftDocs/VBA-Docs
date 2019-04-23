---
title: DataLabel object (Excel Graph)
keywords: vbagr10.chm131186
f1_keywords:
- vbagr10.chm131186
ms.prod: excel
api_name:
- Excel.DataLabel
ms.assetid: 5f823de1-a4c3-bf48-f2fc-c01aabdb9c4d
ms.date: 04/06/2019
localization_priority: Normal
---


# DataLabel object (Excel Graph)

Represents the data label for the specified point or trendline in a chart. 

For a series, the **DataLabel** object is a member of the **[DataLabels](Excel.datalabels(collection).md)** collection, which contains a **DataLabel** object for each point. 

For a series without definable points (such as an area series), the **DataLabels** collection contains a single **DataLabel** object.


## Remarks

Use **DataLabels** (_index_), where _index_ is the data label's index number, to return a single **DataLabel** object.

Use the **[DataLabel](excel.datalabel-graph-property.md)** property to return the **DataLabel** object for a single point. 

## Example

The following example sets the number format for the fifth data label in series one in the chart.

```vb
myChart.SeriesCollection(1).DataLabels(5).NumberFormat = "0.000"
```

<br/>

The following example turns on the data label for the second point in series one in the chart, and sets the data label text to Saturday.

```vb
With myChart 
 With .SeriesCollection(1).Points(2) 
 .HasDataLabel = True 
 .DataLabel.Text = "Saturday" 
 End With 
End With
```

<br/>

For a trendline, the **DataLabel** property returns the text shown with the trendline. This can be the equation, the R-squared value, or both (if both are showing). The following example sets the trendline text to show only the equation, and then places the data label text in cell A1 on the datasheet.

```vb
With myChart.SeriesCollection(1).Trendlines(1) 
 .DisplayRSquared = False 
 .DisplayEquation = True 
 x = .DataLabel.Text 
End With 
With myChart.Application.DataSheet 
 .Range("A1").Value = x 
End With
```


## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]