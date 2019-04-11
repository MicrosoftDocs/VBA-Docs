---
title: ErrorBars object (Excel Graph)
keywords: vbagr10.chm131213
f1_keywords:
- vbagr10.chm131213
ms.prod: excel
api_name:
- Excel.ErrorBars
ms.assetid: f087bede-5ce2-331f-09e1-4c801f8bca82
ms.date: 04/06/2019
localization_priority: Normal
---


# ErrorBars object (Excel Graph)

Represents the error bars for the specified chart series. Error bars indicate the degree of uncertainty for chart data. Only series in area, bar, column, line, and scatter groups in a 2D chart can have error bars. Only series in scatter groups can have x and y error bars.

This object isn't a collection. There's no object that represents a single error bar; either you have x error bars or y error bars turned on for all points in a series or you have them turned off.

## Remarks

Use the **[ErrorBars](excel.errorbars-graph-property.md)** property to return the **ErrorBars** object. 

The **[ErrorBar](Excel.ErrorBar.md)** method changes the format and type of error bars.

## Example

The following example turns on error bars for series one in _myChart_, and then sets the end style for the error bars.

```vb
myChart.SeriesCollection(1).HasErrorBars = True 
myChart.SeriesCollection(1).ErrorBars.EndStyle = xlNoCap
```

## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]