---
title: Order property (Excel Graph)
keywords: vbagr10.chm5207737
f1_keywords:
- vbagr10.chm5207737
ms.prod: excel
api_name:
- Excel.Order
ms.assetid: aa56d241-870c-c3a9-00da-269fb8c314ea
ms.date: 06/08/2017
localization_priority: Normal
---


# Order property (Excel Graph)

Returns or sets the trendline order (an integer greater than 1) when the trendline type is  **xlPolynomial**. Read/write  **Long**.


## Example

This example sets the order of the first trendline for series one if it's polynomial.


```vb
With myChart.SeriesCollection(1).Trendlines(1) 
 If .Type = xlPolynomial Then .Order = 3 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]