---
title: Series.InvertIfNegative property (Excel)
keywords: vbaxl10.chm578092
f1_keywords:
- vbaxl10.chm578092
ms.prod: excel
api_name:
- Excel.Series.InvertIfNegative
ms.assetid: 06c963ac-6e81-5f45-b8b9-8c61bf0c02b6
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.InvertIfNegative property (Excel)

 **True** if Microsoft Excel inverts the pattern in the item when it corresponds to a negative number. Read/write **Boolean**.


## Syntax

_expression_. `InvertIfNegative`

_expression_ A variable that represents a [Series](Excel.Series-graph-object.md) object.


## Example

This example inverts the pattern for negative values in series one in Chart1. The example should be run on a 2-D column chart.


```vb
Charts("Chart1").SeriesCollection(1).InvertIfNegative = True
```


## See also


[Series Object](Excel.Series(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]