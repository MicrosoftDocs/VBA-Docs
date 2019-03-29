---
title: LegendKey object (Excel)
keywords: vbaxl10.chm589072
f1_keywords:
- vbaxl10.chm589072
ms.prod: excel
api_name:
- Excel.LegendKey
ms.assetid: 2d806a8f-2fed-e6f6-bb76-7339fa692cbb
ms.date: 03/30/2019
localization_priority: Normal
---


# LegendKey object (Excel)

Represents a legend key in a chart legend.


## Remarks

Each legend key is a graphic that visually links a legend entry with its associated series or trendline in the chart. The legend key is linked to its associated series or trendline in such a way that changing the formatting of one simultaneously changes the formatting of the other.


## Example

Use the **[LegendKey](Excel.LegendEntry.LegendKey.md)** property of the **LegendEntry** object to return the **LegendKey** object. 

The following example changes the marker background color for the legend entry at the top of the legend for embedded chart one on the worksheet named **Sheet1**. This simultaneously changes the format of every point in the series associated with this legend entry. The associated series must support data markers.

```vb
Worksheets("sheet1").ChartObjects(1).Chart _ 
 .Legend.LegendEntries(1).LegendKey.MarkerBackgroundColorIndex = 5
```

## Methods

- [ClearFormats](Excel.LegendKey.ClearFormats.md)
- [Delete](Excel.LegendKey.Delete.md)

## Properties

- [Application](Excel.LegendKey.Application.md)
- [Creator](Excel.LegendKey.Creator.md)
- [Format](Excel.LegendKey.Format.md)
- [Height](Excel.LegendKey.Height.md)
- [InvertIfNegative](Excel.LegendKey.InvertIfNegative.md)
- [Left](Excel.LegendKey.Left.md)
- [MarkerBackgroundColor](Excel.LegendKey.MarkerBackgroundColor.md)
- [MarkerBackgroundColorIndex](Excel.LegendKey.MarkerBackgroundColorIndex.md)
- [MarkerForegroundColor](Excel.LegendKey.MarkerForegroundColor.md)
- [MarkerForegroundColorIndex](Excel.LegendKey.MarkerForegroundColorIndex.md)
- [MarkerSize](Excel.LegendKey.MarkerSize.md)
- [MarkerStyle](Excel.LegendKey.MarkerStyle.md)
- [Parent](Excel.LegendKey.Parent.md)
- [PictureType](Excel.LegendKey.PictureType.md)
- [PictureUnit2](Excel.LegendKey.PictureUnit2.md)
- [Shadow](Excel.LegendKey.Shadow.md)
- [Smooth](Excel.LegendKey.Smooth.md)
- [Top](Excel.LegendKey.Top.md)
- [Width](Excel.LegendKey.Width.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]