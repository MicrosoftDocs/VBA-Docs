---
title: Interior property (Excel Graph)
keywords: vbagr10.chm65665
f1_keywords:
- vbagr10.chm65665
ms.prod: excel
api_name:
- Excel.Interior
ms.assetid: 5e1fd240-62f6-bb27-8590-283d295ffc76
ms.date: 04/11/2019
localization_priority: Normal
---


# Interior property (Excel Graph)

Returns an **Interior** object that represents the interior of the specified object. Read-only.

## Syntax

_expression_.**Interior**

_expression_ Required. An expression that returns an **[Interior](Excel.Interior-graph-object.md)** object.

## Example

This example sets the interior color of the chart title.

```vb
myChart.ChartTitle.Interior.ColorIndex = 8
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
