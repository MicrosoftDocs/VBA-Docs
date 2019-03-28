---
title: DataLabel object (Excel)
keywords: vbaxl10.chm581072
f1_keywords:
- vbaxl10.chm581072
ms.prod: excel
api_name:
- Excel.DataLabel
ms.assetid: bb342572-8761-b326-548a-98455172f9a8
ms.date: 03/29/2019
localization_priority: Normal
---


# DataLabel object (Excel)

Represents the data label on a chart point or trendline.


## Remarks

In a series, the **DataLabel** object is a member of the **[DataLabels](Excel.DataLabels(object).md)** collection. The **DataLabels** collection contains a **DataLabel** object for each point. For a series without definable points (such as an area series), the **DataLabels** collection contains a single **DataLabel** object.


## Example

Use **[DataLabels](Excel.Series.DataLabels.md)** (_index_), where _index_ is the data-label index number, to return a single **DataLabel** object. The following example sets the number format for the fifth data label in series one in embedded chart one on worksheet one.

```vb
Worksheets(1).ChartObjects(1).Chart _ 
 .SeriesCollection(1).DataLabels(5).NumberFormat = "0.000"
```

<br/>

Use the **[DataLabel](Excel.Point.DataLabel.md)** property of the **Point** object to return the **DataLabel** object for a single point. The following example turns on the data label for the second point in series one on the chart sheet named **Chart1**, and sets the data label text to **Saturday**.

```vb
With Charts("chart1") 
 With .SeriesCollection(1).Points(2) 
 .HasDataLabel = True 
 .DataLabel.Text = "Saturday" 
 End With 
End With
```

<br/>

On a trendline, the **[DataLabel](Excel.Trendline.DataLabel.md)** property returns the text shown with the trendline. This can be the equation, the R-squared value, or both (if both are showing). The following example sets the trendline text to show only the equation, and then places the data label text in cell A1 on the worksheet named **Sheet1**.

```vb
With Charts("chart1").SeriesCollection(1).Trendlines(1) 
 .DisplayRSquared = False 
 .DisplayEquation = True 
 Worksheets("sheet1").Range("a1").Value = .DataLabel.Text 
End With
```


## Methods

- [Delete](Excel.DataLabel.Delete.md)
- [Select](Excel.DataLabel.Select.md)

## Properties

- [Application](Excel.DataLabel.Application.md)
- [AutoText](Excel.DataLabel.AutoText.md)
- [Caption](Excel.DataLabel.Caption.md)
- [Characters](Excel.DataLabel.Characters.md)
- [Creator](Excel.DataLabel.Creator.md)
- [Format](Excel.DataLabel.Format.md)
- [Formula](Excel.DataLabel.Formula.md)
- [FormulaLocal](Excel.DataLabel.FormulaLocal.md)
- [FormulaR1C1](Excel.DataLabel.FormulaR1C1.md)
- [FormulaR1C1Local](Excel.DataLabel.FormulaR1C1Local.md)
- [Height](Excel.DataLabel.Height.md)
- [HorizontalAlignment](Excel.DataLabel.HorizontalAlignment.md)
- [Left](Excel.DataLabel.Left.md)
- [Name](Excel.DataLabel.Name.md)
- [NumberFormat](Excel.DataLabel.NumberFormat.md)
- [NumberFormatLinked](Excel.DataLabel.NumberFormatLinked.md)
- [NumberFormatLocal](Excel.DataLabel.NumberFormatLocal.md)
- [Orientation](Excel.DataLabel.Orientation.md)
- [Parent](Excel.DataLabel.Parent.md)
- [Position](Excel.DataLabel.Position.md)
- [ReadingOrder](Excel.DataLabel.ReadingOrder.md)
- [Separator](Excel.DataLabel.Separator.md)
- [Shadow](Excel.DataLabel.Shadow.md)
- [ShowBubbleSize](Excel.DataLabel.ShowBubbleSize.md)
- [ShowCategoryName](Excel.DataLabel.ShowCategoryName.md)
- [ShowLegendKey](Excel.DataLabel.ShowLegendKey.md)
- [ShowPercentage](Excel.DataLabel.ShowPercentage.md)
- [ShowRange](Excel.datalabel.showrange.md)
- [ShowSeriesName](Excel.DataLabel.ShowSeriesName.md)
- [ShowValue](Excel.DataLabel.ShowValue.md)
- [Text](Excel.DataLabel.Text.md)
- [Top](Excel.DataLabel.Top.md)
- [VerticalAlignment](Excel.DataLabel.VerticalAlignment.md)
- [Width](Excel.DataLabel.Width.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
