---
title: Trendline.Type property (Excel)
keywords: vbaxl10.chm594090
f1_keywords:
- vbaxl10.chm594090
ms.prod: excel
api_name:
- Excel.Trendline.Type
ms.assetid: c07c060c-0512-72a7-c219-d12ea6b629fc
ms.date: 05/18/2019
localization_priority: Normal
---


# Trendline.Type property (Excel)

Returns or sets an **[XlTrendlineType](Excel.XlTrendlineType.md)** value that represents the trendline type.


## Syntax

_expression_.**Type**

_expression_ A variable that represents a **[Trendline](Excel.Trendline(object).md)** object.


## Example

This example changes the trendline type for the first series in embedded chart one on worksheet one. If the series has no trendline, this example fails.

```vb
Worksheets(1).ChartObjects(1).Chart _ 
 .SeriesCollection(1).Trendlines(1).Type = xlMovingAvg
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]