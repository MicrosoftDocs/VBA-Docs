---
title: Intercept property (Excel Graph)
keywords: vbagr10.chm65722
f1_keywords:
- vbagr10.chm65722
ms.prod: excel
api_name:
- Excel.Intercept
ms.assetid: 9c7c4193-8f9d-0f33-74c7-055a9124320e
ms.date: 04/11/2019
localization_priority: Normal
---


# Intercept property (Excel Graph)

Returns or sets the point where the trendline crosses the value axis. Read/write **Double**.

## Syntax

_expression_.**Intercept**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Remarks

Setting this property sets the **[InterceptIsAuto](Excel.InterceptIsAuto.md)** property to **False**.


## Example

This example sets trendline one to cross the value axis at 5. The example should be run on a 2D column chart that contains a single series with a trendline.

```vb
myChart.SeriesCollection(1).Trendlines(1).Intercept = 5
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]