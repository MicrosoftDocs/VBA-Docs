---
title: Series.ErrorBars property (Excel)
keywords: vbaxl10.chm578082
f1_keywords:
- vbaxl10.chm578082
ms.prod: excel
api_name:
- Excel.Series.ErrorBars
ms.assetid: 1a607e6f-e70a-e39c-4cc3-6060eb64e654
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.ErrorBars property (Excel)

Returns an  **[ErrorBars](Excel.ErrorBars(object).md)** object that represents the error bars for the series. Read-only.


## Syntax

_expression_. `ErrorBars`

_expression_ A variable that represents a [Series](Excel.Series-graph-object.md) object.


## Example

This example sets the error bar color for series one in Chart1. The example should be run on a 2-D line chart that has error bars for series one.


```vb
With Charts("Chart1").SeriesCollection(1) 
 .ErrorBars.Border.ColorIndex = 8 
End With
```


## See also


[Series Object](Excel.Series(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]