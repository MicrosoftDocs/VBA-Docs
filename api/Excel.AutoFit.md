---
title: AutoFit method (Excel Graph)
keywords: vbagr10.chm65773
f1_keywords:
- vbagr10.chm65773
ms.prod: excel
api_name:
- Excel.AutoFit
ms.assetid: 45dea7dd-7695-1f72-9bf7-9ab4cbbd74ec
ms.date: 04/06/2019
localization_priority: Normal
---


# AutoFit method (Excel Graph)

Changes the width of the columns in the specified range to achieve the best fit.

## Syntax

_expression_.**AutoFit**

_expression_ Required. An expression that returns a **[Range](excel.range-graph-object.md)** object. 


## Remarks

Must be a row or a range of rows, or a column or a range of columns. Otherwise, this method causes an error.

One unit of column width is equal to the width of one character in the Normal style.


## Example

This example changes the width of columns A through I on the datasheet to achieve the best fit.

```vb
myChart.Application.DataSheet.Columns("A:I").AutoFit
```

<br/>

This example changes the width of columns A through E on the datasheet to achieve the best fit, based only on the contents of cells A1:E1.

```vb
myChart.Application.DataSheet.Range("A1:E1").Columns.AutoFit
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
