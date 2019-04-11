---
title: InterceptIsAuto property (Excel Graph)
keywords: vbagr10.chm65723
f1_keywords:
- vbagr10.chm65723
ms.prod: excel
api_name:
- Excel.InterceptIsAuto
ms.assetid: fd5b2155-8b45-8a67-19c9-8a18a4d3f6f3
ms.date: 04/11/2019
localization_priority: Normal
---


# InterceptIsAuto property (Excel Graph)

**True** if the point where the trendline crosses the value axis is automatically determined by the regression. Read/write **Boolean**.

## Syntax

_expression_.**InterceptIsAuto**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Remarks

Setting the **[Intercept](Excel.Intercept.md)** property sets this property to **False**.


## Example

This example sets Graph to automatically determine the trendline intercept point. The example should be run on a 2D column chart that contains a single series with a trendline.


```vb
myChart.SeriesCollection(1).Trendlines(1) _ 
 .InterceptIsAuto = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]