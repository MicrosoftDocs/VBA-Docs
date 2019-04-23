---
title: ApplyCustomType method (Excel Graph)
keywords: vbagr10.chm66937
f1_keywords:
- vbagr10.chm66937
ms.prod: excel
api_name:
- Excel.ApplyCustomType
ms.assetid: 5385d195-96ce-bdd3-e84d-596fd4236904
ms.date: 04/06/2019
localization_priority: Normal
---


# ApplyCustomType method (Excel Graph)

The **ApplyCustomType** method as it applies to the **Series** and **Chart** objects.

## Series object

Applies a standard or custom chart type to a series.

### Syntax

_expression_.**ApplyCustomType** (_ChartType_)

_expression_ Required. An expression that returns a **[Series](excel.series-graph-object.md)** object.

### Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_ChartType_ |Required |**[XlChartType](excel.xlcharttype.md)** |A standard chart type. Can be one of the **XlChartType** constants. |


## Chart object

Applies a standard or custom chart type to a chart.

### Syntax

_expression_.**ApplyCustomType** (_ChartType_, _TypeName_)

_expression_ Required. An expression that returns a **[Chart](excel.chart-graph-object.md)** object.

### Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_ChartType_ |Required |**[XlChartType](excel.xlcharttype.md)** |A standard chart type. Can be one of the **XlChartType** constants.|
|_TypeName_|Optional |**Variant**|A **String** naming the custom chart type when _ChartType_ specifies a custom chart gallery.|

## Example

This example applies the line with the markers chart type.

```vb
myChart.ApplyCustomType xlLineMarkers
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]