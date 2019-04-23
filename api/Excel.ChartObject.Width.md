---
title: ChartObject.Width property (Excel)
keywords: vbaxl10.chm494095
f1_keywords:
- vbaxl10.chm494095
ms.prod: excel
api_name:
- Excel.ChartObject.Width
ms.assetid: ebe9523f-2777-fd27-a29e-c378355c3c18
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartObject.Width property (Excel)

Returns or sets a **Double** value that represents the width, in [points](../language/glossary/vbe-glossary.md#point), of the object.


## Syntax

_expression_.**Width**

_expression_ A variable that represents a **[ChartObject](Excel.ChartObject.md)** object.


## Example

This example sets the width of the embedded chart.

```vb
Worksheets("Sheet1").ChartObjects(1).Width = 360
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]