---
title: DataLabels object (Excel)
keywords: vbaxl10.chm583072
f1_keywords:
- vbaxl10.chm583072
ms.prod: excel
api_name:
- Excel.DataLabels
ms.assetid: 3d79271e-c702-e785-6984-d838d060a8c5
ms.date: 03/29/2019
localization_priority: Normal
---


# DataLabels object (Excel)

A collection of all the **[DataLabel](Excel.DataLabel(object).md)** objects for the specified series.


## Remarks

Each **DataLabel** object represents a data label for a point or trendline. For a series without definable points (such as an area series), the **DataLabels** collection contains a single data label.


## Example

Use the **[DataLabels](Excel.Series.DataLabels.md)** method of the **Series** object to return the **DataLabels** collection. The following example sets the number format for data labels on series one on chart sheet one.

```vb
With Charts(1).SeriesCollection(1) 
 .HasDataLabels = True 
 .DataLabels.NumberFormat = "##.##" 
End With
```

<br/>

Use **DataLabels** (_index_), where _index_ is the data-label index number, to return a single **DataLabel** object. The following example sets the number format for the fifth data label in series one in embedded chart one on worksheet one.

```vb
Worksheets(1).ChartObjects(1).Chart _ 
 .SeriesCollection(1).DataLabels(5).NumberFormat = "0.000"
```


## Methods

- [Delete](Excel.DataLabels.Delete.md)
- [Item](Excel.DataLabels.Item.md)
- [Propagate](Excel.datalabels.propagate.md)
- [Select](Excel.DataLabels.Select.md)

## Properties

- [Application](Excel.DataLabels.Application.md)
- [AutoText](Excel.DataLabels.AutoText.md)
- [Count](Excel.DataLabels.Count.md)
- [Creator](Excel.DataLabels.Creator.md)
- [Format](Excel.DataLabels.Format.md)
- [HorizontalAlignment](Excel.DataLabels.HorizontalAlignment.md)
- [Name](Excel.DataLabels.Name.md)
- [NumberFormat](Excel.DataLabels.NumberFormat.md)
- [NumberFormatLinked](Excel.DataLabels.NumberFormatLinked.md)
- [NumberFormatLocal](Excel.DataLabels.NumberFormatLocal.md)
- [Orientation](Excel.DataLabels.Orientation.md)
- [Parent](Excel.DataLabels.Parent.md)
- [Position](Excel.DataLabels.Position.md)
- [ReadingOrder](Excel.DataLabels.ReadingOrder.md)
- [Separator](Excel.DataLabels.Separator.md)
- [Shadow](Excel.DataLabels.Shadow.md)
- [ShowBubbleSize](Excel.DataLabels.ShowBubbleSize.md)
- [ShowCategoryName](Excel.DataLabels.ShowCategoryName.md)
- [ShowLegendKey](Excel.DataLabels.ShowLegendKey.md)
- [ShowPercentage](Excel.DataLabels.ShowPercentage.md)
- [ShowRange](Excel.datalabels.showrange.md)
- [ShowSeriesName](Excel.DataLabels.ShowSeriesName.md)
- [ShowValue](Excel.DataLabels.ShowValue.md)
- [VerticalAlignment](Excel.DataLabels.VerticalAlignment.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
