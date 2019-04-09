---
title: DataSheet property (Excel Graph)
keywords: vbagr10.chm5207292
f1_keywords:
- vbagr10.chm5207292
ms.prod: excel
api_name:
- Excel.DataSheet
ms.assetid: d7ccd394-e9b7-2967-76a4-60e5dda40a84
ms.date: 04/10/2019
localization_priority: Normal
---


# DataSheet property (Excel Graph)

Returns the **DataSheet** object. Read-only.

## Syntax

_expression_.**DataSheet**

_expression_ Required. An expression that returns a **[DataSheet](Excel.DataSheet-graph-object.md)** object.

## Example

This example sets the value of cell A1 on the datasheet to 3.14159.

```vb
With myChart.Application 
 .DataSheet.Range("A1").Value = 3.14159 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]