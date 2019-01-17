---
title: Chart.PlotBy property (Excel)
keywords: vbaxl10.chm149155
f1_keywords:
- vbaxl10.chm149155
ms.prod: excel
api_name:
- Excel.Chart.PlotBy
ms.assetid: 69ff0fbe-7954-6808-68fa-cc92b2851dd8
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.PlotBy property (Excel)

Returns or sets the way columns or rows are used as data series on the chart. Can be one of the following  **[xlRowCol](Excel.XlRowCol.md)** constants: **xlColumns** or **xlRows**. Read/write **Long**.


## Syntax

_expression_. `PlotBy`

_expression_ A variable that represents a [Chart](Excel.Chart-graph-object.md) object.


## Remarks

For PivotChart reports, this property is read-only and always returns  **xlColumns**.


## Example

This example causes the embedded chart to plot data by columns.


```vb
Worksheets(1).ChartObjects(1).Chart.PlotBy = xlColumns
```


## See also


[SparklineGroup Object](Excel.SparklineGroup.md)
[Chart Object](Excel.Chart(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]