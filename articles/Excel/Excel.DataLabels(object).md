---
title: DataLabels Object (Excel)
keywords: vbaxl10.chm583072
f1_keywords:
- vbaxl10.chm583072
ms.prod: excel
api_name:
- Excel.DataLabels
ms.assetid: 3d79271e-c702-e785-6984-d838d060a8c5
ms.date: 06/08/2017
---


# DataLabels Object (Excel)

A collection of all the  **[DataLabel](Excel.DataLabel(objec).md)** objects for the specified series.


## Remarks

 Each **DataLabel** object represents a data label for a point or trendline. For a series without definable points (such as an area series), the **DataLabels** collection contains a single data label.


## Example

Use the  **[DataLabels](Excel.Series.DataLabels.md)** method to return the **DataLabels** collection. The following example sets the number format for data labels on series one on chart sheet one.


```
With Charts(1).SeriesCollection(1) 
 .HasDataLabels = True 
 .DataLabels.NumberFormat = "##.##" 
End With
```

Use  **DataLabels** ( _index_ ), where _index_ is the data-label index number, to return a single **DataLabel** object. The following example sets the number format for the fifth data label in series one in embedded chart one on worksheet one.




```
Worksheets(1).ChartObjects(1).Chart _ 
 .SeriesCollection(1).DataLabels(5).NumberFormat = "0.000"
```


## Methods



|**Name**|
|:-----|
|[Delete](Excel.DataLabels.Delete.md)|
|[Item](Excel.DataLabels.Item.md)|
|[Propagate](Excel.datalabels.propagate.md)|
|[Select](Excel.DataLabels.Select.md)|

## Properties



|**Name**|
|:-----|
|[Application](Excel.DataLabels.Application.md)|
|[AutoText](Excel.DataLabels.AutoText.md)|
|[Count](Excel.DataLabels.Count.md)|
|[Creator](Excel.DataLabels.Creator.md)|
|[Format](Excel.DataLabels.Format.md)|
|[HorizontalAlignment](Excel.DataLabels.HorizontalAlignment.md)|
|[Name](Excel.DataLabels.Name.md)|
|[NumberFormat](Excel.DataLabels.NumberFormat.md)|
|[NumberFormatLinked](Excel.DataLabels.NumberFormatLinked.md)|
|[NumberFormatLocal](Excel.DataLabels.NumberFormatLocal.md)|
|[Orientation](Excel.DataLabels.Orientation.md)|
|[Parent](Excel.DataLabels.Parent.md)|
|[Position](Excel.DataLabels.Position.md)|
|[ReadingOrder](Excel.DataLabels.ReadingOrder.md)|
|[Separator](Excel.DataLabels.Separator.md)|
|[Shadow](Excel.DataLabels.Shadow.md)|
|[ShowBubbleSize](Excel.DataLabels.ShowBubbleSize.md)|
|[ShowCategoryName](Excel.DataLabels.ShowCategoryName.md)|
|[ShowLegendKey](Excel.DataLabels.ShowLegendKey.md)|
|[ShowPercentage](Excel.DataLabels.ShowPercentage.md)|
|[ShowRange](Excel.datalabels.showrange.md)|
|[ShowSeriesName](Excel.DataLabels.ShowSeriesName.md)|
|[ShowValue](Excel.DataLabels.ShowValue.md)|
|[VerticalAlignment](Excel.DataLabels.VerticalAlignment.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
