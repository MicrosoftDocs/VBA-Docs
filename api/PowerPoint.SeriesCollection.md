---
title: SeriesCollection object (PowerPoint)
keywords: vbapp10.chm717000
f1_keywords:
- vbapp10.chm717000
ms.prod: powerpoint
api_name:
- PowerPoint.SeriesCollection
ms.assetid: 6277f9e0-0198-0773-9c54-f2d009c0ba7a
ms.date: 06/08/2017
localization_priority: Normal
---


# SeriesCollection object (PowerPoint)

Represents a collection of all the  **[Series](PowerPoint.Series.md)** objects in the specified chart or chart group.


## Remarks

Use the  **[SeriesCollection](PowerPoint.Chart.SeriesCollection.md)** method to return the **SeriesCollection** collection.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

 Use the **[Extend](PowerPoint.SeriesCollection.Extend.md)** method to extend an existing series. The following example adds the data in cells C6:C10 in the chart's worksheet to an existing series in the series collection of the chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection.Extend "='Sheet1'!$C$6:$C$10"

    End If

End With
```




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use the  **[Add](PowerPoint.SeriesCollection.Add.md)** method to create a new series and add it to the chart. The following example adds the data from cells D1:D5 in the chart's worksheet as a new series to the chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection.Add "='Sheet1'!$D$1:$D$5"

    End If

End With
```




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

Use  **SeriesCollection** (_index_), where _index_ is the series index number or name, to return a single **Series** object. The following example sets the color of the interior for the first series in embedded chart one the active document.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.SeriesCollection(1).Interior.Color = RGB(255, 0, 0)

    End If

End With
```


## Methods



|Name|
|:-----|
|[Add](PowerPoint.SeriesCollection.Add.md)|
|[Extend](PowerPoint.SeriesCollection.Extend.md)|
|[Item](PowerPoint.SeriesCollection.Item.md)|
|[NewSeries](PowerPoint.SeriesCollection.NewSeries.md)|

## Properties



|Name|
|:-----|
|[Application](PowerPoint.SeriesCollection.Application.md)|
|[Count](PowerPoint.SeriesCollection.Count.md)|
|[Creator](PowerPoint.SeriesCollection.Creator.md)|
|[Parent](PowerPoint.SeriesCollection.Parent.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]