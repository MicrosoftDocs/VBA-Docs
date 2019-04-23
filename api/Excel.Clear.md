---
title: Clear method (Excel Graph)
keywords: vbagr10.chm65647
f1_keywords:
- vbagr10.chm65647
ms.prod: excel
api_name:
- Excel.Clear
ms.assetid: f77c2fc0-6ec4-7345-0e5c-7b8dd4cd1a90
ms.date: 04/06/2019
localization_priority: Normal
---


# Clear method (Excel Graph)

The **Clear** method as it applies to the **ChartArea**, **Legend**, and **Range** objects.

## ChartArea and Legend objects

Clears the entire chart area.

### Syntax

_expression_.**Clear**

_expression_ Required. An expression that returns a **[ChartArea](excel.chartarea-graph-object.md)** or **[Legend](excel.legend-graph-object.md)** object.


## Range object

Clears the entire range.

### Syntax

_expression_.**Clear**

_expression_ Required. An expression that returns a **[Range](excel.range-graph-object.md)** object. 

## Example

This example clears the formulas and formatting in cells A1:G37 on the datasheet.

```vb
myChart.Application.DataSheet.Range("A1:G37").Clear
```

<br/>

This example clears the chart area (the chart data and formatting) of Chart1.

```vb
myChart.ChartArea.Clear
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]