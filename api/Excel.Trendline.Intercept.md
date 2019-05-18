---
title: Trendline.Intercept property (Excel)
keywords: vbaxl10.chm594083
f1_keywords:
- vbaxl10.chm594083
ms.prod: excel
api_name:
- Excel.Trendline.Intercept
ms.assetid: a3a1b427-2da2-4409-5488-20a1eb0ceb94
ms.date: 05/18/2019
localization_priority: Normal
---


# Trendline.Intercept property (Excel)

Returns or sets the point where the trendline crosses the value axis. Read/write **Double**.


## Syntax

_expression_.**Intercept**

 _expression_ An expression that returns a **[Trendline](Excel.Trendline(object).md)** object.


## Return value

Double


## Remarks

Setting this property sets the **[InterceptIsAuto](Excel.Trendline.InterceptIsAuto.md)** property to **False**.


## Example

This example sets trendline one on Chart1 to cross the value axis at 5. The example should be run on a 2D column chart that contains a single series with a trendline.

```vb
Charts("Chart1").SeriesCollection(1).Trendlines(1).Intercept = 5
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]