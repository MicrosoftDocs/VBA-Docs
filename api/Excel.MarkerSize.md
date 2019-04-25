---
title: MarkerSize property (Excel Graph)
keywords: vbagr10.chm65767
f1_keywords:
- vbagr10.chm65767
ms.prod: excel
api_name:
- Excel.MarkerSize
ms.assetid: 8af0ad97-9291-8bee-b896-283b76f8a882
ms.date: 04/11/2019
localization_priority: Normal
---


# MarkerSize property (Excel Graph)

Returns or sets the data-marker size, in [points](../language/glossary/vbe-glossary.md#point). Read/write **Long**.


## Syntax

_expression_.**MarkerSize**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets the data-marker size for all data markers in series one.

```vb
MyChart.SeriesCollection(1).MarkerSize = 10
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]