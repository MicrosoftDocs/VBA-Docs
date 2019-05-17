---
title: Trendline.InterceptIsAuto property (Excel)
keywords: vbaxl10.chm594084
f1_keywords:
- vbaxl10.chm594084
ms.prod: excel
api_name:
- Excel.Trendline.InterceptIsAuto
ms.assetid: ec5ea945-59d7-3ec2-42cd-95c7031880e8
ms.date: 05/18/2019
localization_priority: Normal
---


# Trendline.InterceptIsAuto property (Excel)

**True** if the point where the trendline crosses the value axis is automatically determined by the regression. Read/write **Boolean**.


## Syntax

_expression_.**InterceptIsAuto**

_expression_ A variable that represents a **[Trendline](Excel.Trendline(object).md)** object.


## Remarks

Setting the **[Intercept](Excel.Trendline.Intercept.md)** property sets this property to **False**.


## Example

This example sets Microsoft Excel to automatically determine the trendline intercept point for Chart1. The example should be run on a 2D column chart that contains a single series with a trendline.

```vb
Charts("Chart1").SeriesCollection(1).Trendlines(1) _ 
 .InterceptIsAuto = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]