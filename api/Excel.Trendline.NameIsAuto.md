---
title: Trendline.NameIsAuto property (Excel)
keywords: vbaxl10.chm594086
f1_keywords:
- vbaxl10.chm594086
ms.prod: excel
api_name:
- Excel.Trendline.NameIsAuto
ms.assetid: 4e14cc52-a9f5-3dda-8be9-7afd97d79583
ms.date: 05/18/2019
localization_priority: Normal
---


# Trendline.NameIsAuto property (Excel)

**True** if Microsoft Excel automatically determines the name of the trendline. Read/write **Boolean**.


## Syntax

_expression_.**NameIsAuto**

_expression_ A variable that represents a **[Trendline](Excel.Trendline(object).md)** object.


## Example

This example sets Microsoft Excel to automatically determine the name for trendline one on Chart1. The example should be run on a 2D column chart that contains a single series with a trendline.

```vb
Charts("Chart1").SeriesCollection(1) _ 
 .Trendlines(1).NameIsAuto = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]