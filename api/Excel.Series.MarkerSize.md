---
title: Series.MarkerSize property (Excel)
keywords: vbaxl10.chm578097
f1_keywords:
- vbaxl10.chm578097
ms.prod: excel
api_name:
- Excel.Series.MarkerSize
ms.assetid: d1e499ae-d59c-3493-c741-9607c3c27a17
ms.date: 05/11/2019
localization_priority: Normal
---


# Series.MarkerSize property (Excel)

Returns or sets the data-marker size, in [points](../language/glossary/vbe-glossary.md#point). Can be a value from 2 through 72. Read/write **Long**.


## Syntax

_expression_.**MarkerSize**

_expression_ A variable that represents a **[Series](Excel.Series(object).md)** object.


## Example

This example sets the data-marker size for all data markers on series one.

```vb
Worksheets(1).ChartObjects(1).Chart _ 
 .SeriesCollection(1).MarkerSize = 10
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
