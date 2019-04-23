---
title: ClearContents method (Excel Graph)
keywords: vbagr10.chm65649
f1_keywords:
- vbagr10.chm65649
ms.prod: excel
api_name:
- Excel.ClearContents
ms.assetid: 8bf70623-e644-e45e-1b1e-565fe6acd223
ms.date: 04/06/2019
localization_priority: Normal
---


# ClearContents method (Excel Graph)

The **ClearContents** method as it applies to the **ChartArea** and **Range** objects.

## ChartArea object

Clears the data from a chart but leaves the formatting.

### Syntax

_expression_.**ClearContents**

_expression_ Required. An expression that returns a **[ChartArea](excel.chartarea-graph-object.md)** object.



## Range object

Clears the formulas from the range.

### Syntax

_expression_.**ClearContents**

_expression_ Required. An expression that returns a **[Range](excel.range-graph-object.md)** object. 

## Example

This example clears the formulas from cells A1:G37 on the datasheet but leaves the formatting intact.

```vb
myChart.Application.DataSheet.Range("A1:G37").ClearContents
```

<br/>

This example clears the chart data from a chart but leaves the formatting intact.

```vb
myChart.ChartArea.ClearContents
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]