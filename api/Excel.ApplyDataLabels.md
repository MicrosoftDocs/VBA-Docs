---
title: ApplyDataLabels method (Excel Graph)
keywords: vbagr10.chm67458
f1_keywords:
- vbagr10.chm67458
ms.prod: excel
api_name:
- Excel.ApplyDataLabels
ms.assetid: 1750d716-66f8-fe4e-8023-fbcfcc5c5ff5
ms.date: 04/06/2019
localization_priority: Normal
---


# ApplyDataLabels method (Excel Graph)

The **ApplyDataLabels** method as it applies to the **Chart**, **Point**, and **Series** objects.

## Chart object

Applies data labels to a point, a series, or all the series in a chart.

### Syntax

_expression_.**ApplyDataLabels** (_Type_, _LegendKey_, _AutoText_, _HasLeaderLines_)

_expression_ Required. An expression that returns a **[Chart](Excel.Chart-graph-object.md)** object.
 
### Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Type_|Optional| **[XlDataLabelsType](excel.xldatalabelstype.md)**|The data label type. Can be one of the **XlDataLabelsType** constants.|
|_LegendKey_ |Optional |**Variant**|**True** to show the legend key next to the point. The default value is **False**.|
| _AutoText_ |Optional |**Variant**|**True** if the object automatically generates appropriate text based on content.|
| _HasLeaderLines_ |Optional |**Variant**|**True** if the series has leader lines.|

## Point and Series objects

Applies data labels to a point, a series, or all the series in a chart.

### Syntax

_expression_.**ApplyDataLabels** (_Type_, _LegendKey_, _AutoText_, _HasLeaderLines_, _ShowSeriesName_, _ShowCategoryName_, _ShowValue_, _ShowPercentage_, _ShowBubbleSize_, _Separator_)

_expression_ Required. An expression that returns a **[Point](excel.point-graph-object.md)** or **[Series](excel.series-graph-object.md)** object.


### Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Type_|Optional| **[XlDataLabelsType](excel.xldatalabelstype.md)**|The data label type. Can be one of the **XlDataLabelsType** constants.|
|_LegendKey_ |Optional |**Variant**|**True** to show the legend key next to the point. The default value is **False**.|
| _AutoText_ |Optional |**Variant**|**True** if the object automatically generates appropriate text based on content.|
| _HasLeaderLines_ |Optional |**Variant**|**True** if the series has leader lines.|
| _ShowSeriesName_ |Optional |**Variant**|The series name for the data label.|
| _ShowCategoryName_ |Optional |**Variant**|The category name for the data label.|
| _ShowValue_ |Optional |**Variant**|The value for the data label.|
| _ShowPercentage_ |Optional |**Variant**|The percentage for the data label.|
| _ShowBubbleSize_ |Optional |**Variant**|The bubble size for the data label.|
| _Separator_ |Optional |**Variant**|The separator for the data label.|

## Example

This example applies category labels to series one.

```vb
myChart.SeriesCollection(1). _ 
 ApplyDataLabels Type:=xlDataLabelsShowLabel
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]