---
title: ErrorBars object (Excel)
keywords: vbaxl10.chm623072
f1_keywords:
- vbaxl10.chm623072
ms.prod: excel
api_name:
- Excel.ErrorBars
ms.assetid: 646de974-bf6f-99c8-20dd-9ca514b7a304
ms.date: 03/29/2019
localization_priority: Normal
---


# ErrorBars object (Excel)

Represents the error bars in a chart series.


## Remarks

Error bars indicate the degree of uncertainty for chart data. Only series in area, bar, column, line, and scatter groups on a 2D chart can have error bars. Only series in scatter groups can have x and y error bars. This object isn't a collection. There's no object that represents a single error bar; you either have x error bars or y error bars turned on for all points in a series or you have them turned off.

The **[ErrorBar](Excel.Series.ErrorBar.md)** method of the **Series** object changes the error bar format and type.


## Example

Use the **[ErrorBars](Excel.Series.ErrorBars.md)** property of the **Series** object to return the **ErrorBars** object. The following example turns on error bars for series one in embedded chart one, and then sets the end style for the error bars.

```vb
Worksheets("sheet1").ChartObjects(1).Activate 
ActiveChart.SeriesCollection(1).HasErrorBars = True 
ActiveChart.SeriesCollection(1).ErrorBars.EndStyle = xlNoCap
```

## Methods

- [ClearFormats](Excel.ErrorBars.ClearFormats.md)
- [Delete](Excel.ErrorBars.Delete.md)
- [Select](Excel.ErrorBars.Select.md)

## Properties

- [Application](Excel.ErrorBars.Application.md)
- [Border](Excel.ErrorBars.Border.md)
- [Creator](Excel.ErrorBars.Creator.md)
- [EndStyle](Excel.ErrorBars.EndStyle.md)
- [Format](Excel.ErrorBars.Format.md)
- [Name](Excel.ErrorBars.Name.md)
- [Parent](Excel.ErrorBars.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]