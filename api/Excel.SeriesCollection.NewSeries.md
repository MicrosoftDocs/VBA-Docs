---
title: SeriesCollection.NewSeries method (Excel)
keywords: vbaxl10.chm580080
f1_keywords:
- vbaxl10.chm580080
ms.prod: excel
api_name:
- Excel.SeriesCollection.NewSeries
ms.assetid: 1d63ff48-d4ec-ce76-42bb-c5923251bd69
ms.date: 06/08/2017
localization_priority: Normal
---


# SeriesCollection.NewSeries method (Excel)

Creates a new series. Returns a  **[Series](Excel.Series(object).md)** object that represents the new series.


## Syntax

_expression_.**NewSeries**

_expression_ A variable that represents a [SeriesCollection](Excel.SeriesCollection.md) object.


## Return value

Series


## Remarks

This method isn't available for PivotChart reports.


## Example

This example adds a new series to chart one.


```vb
Set ns = Charts(1).SeriesCollection.NewSeries
```


## See also


[SeriesCollection Object](Excel.SeriesCollection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
