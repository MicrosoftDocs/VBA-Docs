---
title: ChartArea object (Excel)
keywords: vbaxl10.chm619072
f1_keywords:
- vbaxl10.chm619072
ms.prod: excel
api_name:
- Excel.ChartArea
ms.assetid: 883423b5-7689-b164-c0a3-8dab049b5d9e
ms.date: 03/29/2019
localization_priority: Normal
---


# ChartArea object (Excel)

Represents the chart area of a chart. 


## Remarks

The chart area includes everything, including the plot area. However, the plot area has its own fill, so filling the plot area does not fill the chart area.

For information about formatting the plot area, see the **[PlotArea](Excel.PlotArea(object).md)** object.

Use the **[ChartArea](excel.chart.chartarea.md)** property of the **Chart** object to return the **ChartArea** object.


## Example

The following example turns off the border for the chart area in embedded chart 1 on the worksheet named **Sheet1**.

```vb
Worksheets("Sheet1").ChartObjects(1).Chart. _ 
 ChartArea.Format.Line.Visible = False
```


## Methods

- [Clear](Excel.ChartArea.Clear.md)
- [ClearContents](Excel.ChartArea.ClearContents.md)
- [ClearFormats](Excel.ChartArea.ClearFormats.md)
- [Copy](Excel.ChartArea.Copy.md)
- [Select](Excel.ChartArea.Select.md)

## Properties

- [Application](Excel.ChartArea.Application.md)
- [Creator](Excel.ChartArea.Creator.md)
- [Format](Excel.ChartArea.Format.md)
- [Height](Excel.ChartArea.Height.md)
- [Left](Excel.ChartArea.Left.md)
- [Name](Excel.ChartArea.Name.md)
- [Parent](Excel.ChartArea.Parent.md)
- [RoundedCorners](Excel.ChartArea.RoundedCorners.md)
- [Shadow](Excel.ChartArea.Shadow.md)
- [Top](Excel.ChartArea.Top.md)
- [Width](Excel.ChartArea.Width.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
